import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPComponentLoader } from '@microsoft/sp-loader';

// --- DOMPurify: safe import + minimal typing ---
import * as DomPurifyNs from 'dompurify';
interface IDomPurify { sanitize(input: string, config?: unknown): string; }
const DOMPURIFY: IDomPurify =
  ((DomPurifyNs as unknown as { default?: IDomPurify }).default ??
   (DomPurifyNs as unknown as IDomPurify));

// Type-only import for the PnP Code Editor (erased at emit)
import type * as PnpCodeEditorNS from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';

export interface IScriptEditorWebPartProps {
  // PnP-like UX
  editTitle: string;            // Title to show in edit mode
  keepPadding: boolean;         // When false, remove top/bottom padding around the part
  enableClassicContext: boolean;// When true, populate window._spPageContextInfo

  // Security/behavior
  markup: string;
  useIframe: boolean;
  allowInlineScript: boolean;
  domSanitize: boolean;
  cspStrict: boolean;
  /** Comma-separated allowlist. e.g., "self, *.sharepoint.com, cdn.contoso.com" */
  allowedDomains: string;
}

export default class ScriptEditorWebPart
  extends BaseClientSideWebPart<IScriptEditorWebPartProps> {

  private _iframe: HTMLIFrameElement | null = null;

  // Background lazy-load for PnP editor (no blocking)
  private _pnpLoaded = false;
  private _pnpLoadPromise: Promise<void> | null = null;
  private _Pnp: typeof PnpCodeEditorNS | null = null;

  public async onInit(): Promise<void> {
    // Optional classic context injection at start
    if (this.properties.enableClassicContext) this._ensureClassicContext();
    return super.onInit();
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
  public get disableReactivePropertyChanges(): boolean { return true; }

  public render(): void {
    try {
      // Container padding adjust like Modern Script Editor
      this._applyHostPadding();

      // Title overlay in edit mode (purely cosmetic)
      this._renderEditTitleOverlay();

      const html: string = this.properties.markup || '';
      const useIframe: boolean = !!this.properties.useIframe;

      if (!useIframe && this._iframe) {
        try { this._iframe.remove(); } catch { /* ignore */ }
        this._iframe = null;
      }
      if (!useIframe) this.domElement.innerHTML = '';
      else if (this._iframe && !this._iframe.isConnected) this._iframe = null;

      if (useIframe) this._renderInIframe(html);
      else this._renderInline(html);

      // Keep _spPageContextInfo aligned if toggled later
      if (this.properties.enableClassicContext) this._ensureClassicContext();

    } catch (e: unknown) {
      this._displayError(e);
      // eslint-disable-next-line no-console
      console.error('ScriptEditor render error:', e);
    }
  }

  // -------------------- Padding + Edit title like PnP --------------------

  private _applyHostPadding(): void {
    const host = this._findHostContainer();
    if (!host) return;

    const removePadding = this.properties.keepPadding === false;
    // Only top/bottom like PnP toggle
    host.style.paddingTop = removePadding ? '0' : '';
    host.style.paddingBottom = removePadding ? '0' : '';
  }

  private _findHostContainer(): HTMLElement | null {
    // Try common wrappers; fall back to two levels up
    const closest = (this.domElement.closest('.ControlZone') ||
                     this.domElement.closest('[data-sp-feature-instance-id]')) as HTMLElement | null;
    if (closest) return closest;
    const p = this.domElement.parentElement;
    return (p && p.parentElement) ? p.parentElement as HTMLElement : p as HTMLElement | null;
  }

  private _renderEditTitleOverlay(): void {
    // The overlay is visible only in edit mode, like the PnP sample
    const shouldShow = this.displayMode === DisplayMode.Edit && !!this.properties.editTitle;
    const id = 'msed-title-overlay';
    let overlay = this.domElement.querySelector(`#${id}`) as HTMLDivElement | null;

    if (!shouldShow) {
      if (overlay) overlay.remove();
      return;
    }

    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = id;
      overlay.style.position = 'relative';
      overlay.style.fontFamily = 'Segoe UI, Arial, sans-serif';
      overlay.style.fontSize = '12px';
      overlay.style.color = '#605e5c';
      overlay.style.marginBottom = '6px';
      this.domElement.prepend(overlay);
    }
    overlay.textContent = this.properties.editTitle;
  }

  // -------------------- Classic _spPageContextInfo --------------------

  private _ensureClassicContext(): void {
    const pc = this.context.pageContext;
    (window as unknown as { _spPageContextInfo?: any })._spPageContextInfo = {
      webAbsoluteUrl: pc.web.absoluteUrl,
      siteAbsoluteUrl: pc.site.absoluteUrl,
      userDisplayName: pc.user.displayName,
      userLoginName: pc.user.loginName,
      currentCultureName: pc.cultureInfo.currentCultureName,
      currentUICultureName: pc.cultureInfo.currentUICultureName,
      serverRequestPath: location.pathname
      // Add more if you truly need them
    };
  }

  // -------------------- Error UI --------------------

  private _displayError(e: unknown): void {
    const msg: string =
      typeof e === 'string' ? e :
      (e && (e as { message?: string }).message) ? (e as { message: string }).message :
      JSON.stringify(e);

    this.domElement.innerHTML = `
      <div style="border:1px solid #e0e0e0;border-left:4px solid #d83b01;padding:10px 12px;border-radius:6px;background:#fff8f6;color:#323130;font-family:Segoe UI,Arial,sans-serif;">
        <div style="font-weight:600;margin-bottom:6px;">Script Editor error</div>
        <div>${this._escapeHtml(msg)}</div>
      </div>`;
  }

  private _escapeHtml(s: string): string {
    const map: Record<string, string> = {
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
    };
    return s.replace(/[&<>"']/g, (c) => map[c]);
  }

  // -------------------- Allowlist helpers --------------------

  private _parseAllowlist(): string[] {
    return (this.properties.allowedDomains || '')
      .split(',')
      .map(v => v.trim().toLowerCase())
      .filter(v => v.length > 0);
  }

  private _isUrlAllowed(url: string): boolean {
    try {
      const u = new URL(url, window.location.href);
      const host = u.hostname.toLowerCase();
      const allow = this._parseAllowlist();

      for (const rule of allow) {
        if (rule === 'self' && u.origin === window.location.origin) return true;
        if (rule.startsWith('*.') && host.endsWith(rule.slice(1))) return true;
        if (rule === host || rule === u.origin.toLowerCase()) return true;
      }
      return false;
    } catch {
      return false;
    }
  }

  // -------------------- IFRAME MODE --------------------

  private _renderInIframe(markup: string): void {
    if (!this._iframe) {
      const f = document.createElement('iframe');
      f.style.width = '100%';
      f.style.border = '0';
      f.style.minHeight = '120px';
      f.setAttribute('sandbox', this._sandboxAttr());
      this.domElement.appendChild(f);
      this._iframe = f;
    } else if (!this._iframe.isConnected) {
      this.domElement.appendChild(this._iframe);
    }

    const doc = this._iframe.contentDocument;
    if (!doc) throw new Error('iFrame contentDocument is not available.');

    const head = `
      <meta charset="utf-8">
      ${this._cspMetaTag()}
      <style>html,body{margin:0;padding:0}body{font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#323130}</style>
    `;

    const sanitized: string = this.properties.domSanitize
      ? DOMPURIFY.sanitize(markup, { USE_PROFILES: { html: true } })
      : markup;

    doc.open();
    doc.write(`<html><head>${head}</head><body>${sanitized}</body></html>`);
    doc.close();

    const resize = (): void => {
      try {
        if (!this._iframe) return;
        const h = Math.max(
          doc.body.scrollHeight, doc.documentElement.scrollHeight,
          doc.body.offsetHeight, doc.documentElement.offsetHeight
        );
        this._iframe.style.height = `${Math.max(h, 120)}px`;
      } catch (err) { this._displayError(err as unknown); }
    };
    setTimeout(resize, 50);
    setTimeout(resize, 300);
  }

  private _sandboxAttr(): string {
    return ['allow-scripts', 'allow-same-origin'].join(' ');
  }

  private _cspMetaTag(): string {
    const allow = this._parseAllowlist().filter(d => d !== 'self');
    const scriptSrc: string[] = [`'self'`, ...allow];
    if (this.properties.allowInlineScript && !this.properties.cspStrict) scriptSrc.push(`'unsafe-inline'`);
    const styleSrc = [`'self'`, `'unsafe-inline'`, 'data:'];
    const imgSrc = [`'self'`, 'data:'];
    const connectSrc = [`'self'`, ...allow];

    const csp = [
      `default-src 'self'`,
      `script-src ${scriptSrc.join(' ')}`,
      `style-src ${styleSrc.join(' ')}`,
      `img-src ${imgSrc.join(' ')}`,
      `connect-src ${connectSrc.join(' ')}`
    ].join('; ');
    return `<meta http-equiv="Content-Security-Policy" content="${csp}">`;
  }

  // -------------------- INLINE MODE --------------------

  private _renderInline(markup: string): void {
    const { htmlNoScripts, externalScripts, inlineScripts } = this._extractScripts(markup);

    const safeHtml: string = this.properties.domSanitize
      ? DOMPURIFY.sanitize(htmlNoScripts, { FORBID_TAGS: ['script'] })
      : htmlNoScripts;

    this.domElement.innerHTML = safeHtml;

    const loadExternal = async (): Promise<void> => {
      for (const src of externalScripts) {
        if (!this._isUrlAllowed(src)) {
          this._displayError(`Blocked external script from non-allowed domain: ${src}`);
          continue;
        }
        await SPComponentLoader.loadScript(src);
      }
    };

    const runInline = (): void => {
      if (!this.properties.allowInlineScript) return;
      for (const code of inlineScripts) {
        const s = document.createElement('script');
        s.type = 'text/javascript';
        s.text = code;
        this.domElement.appendChild(s);
      }
    };

    loadExternal().then(runInline).catch((err: unknown) => this._displayError(err));
  }

  private _extractScripts(markup: string): {
    htmlNoScripts: string;
    externalScripts: string[];
    inlineScripts: string[];
  } {
    const parser = new DOMParser();
    const doc = parser.parseFromString(markup, 'text/html');
    const externalScripts: string[] = [];
    const inlineScripts: string[] = [];

    Array.from(doc.getElementsByTagName('script')).forEach((sc: HTMLScriptElement) => {
      const src = sc.getAttribute('src');
      if (src) externalScripts.push(src);
      else inlineScripts.push(sc.text || '');
      sc.parentNode?.removeChild(sc);
    });

    const htmlNoScripts = doc.body ? doc.body.innerHTML : markup;
    return { htmlNoScripts, externalScripts, inlineScripts };
  }

  // -------------------- Property Pane --------------------

  /** Kick off background load; pane opens immediately with fallback */
  private _ensurePnpLoaded(): void {
    if (this._pnpLoaded || this._pnpLoadPromise) return;
    this._pnpLoadPromise = import(
      /* webpackChunkName: 'pnp-property-controls' */
      '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    )
      .then(mod => {
        this._Pnp = mod;
        this._pnpLoaded = true;
        this._pnpLoadPromise = null;
        this.context.propertyPane.refresh(); // swap to rich editor if pane open
      })
      .catch(err => {
        // eslint-disable-next-line no-console
        console.warn('PnP Code Editor failed to load; using fallback.', err);
        this._pnpLoadPromise = null;
      });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this._ensurePnpLoaded();

    const codeField =
      this._Pnp
        ? this._Pnp.PropertyFieldCodeEditor('markup', {
            label: 'Code (HTML/JS/CSS)',
            panelTitle: 'Edit HTML Code',
            initialValue: this.properties.markup,
            onPropertyChange: this.onPropertyPaneFieldChanged,
            properties: this.properties,
            key: 'codeEditorField',
            language: this._Pnp.PropertyFieldCodeEditorLanguages.HTML,
            options: { wrap: true, showPrintMargin: false, tabSize: 2, useWorker: false }
          })
        : PropertyPaneTextField('markup', {
            label: 'Code (HTML/JS/CSS)',
            multiline: true, resizable: true, rows: 12
          });

    return {
      pages: [
        {
          header: { description: 'Script Editor Settings (org-hardened)' },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('editTitle', {
                  label: 'Title to show in edit mode'
                }),
                PropertyPaneToggle('keepPadding', {
                  label: 'Keep padding', onText: 'Keep', offText: 'Remove top/bottom padding'
                }),
                PropertyPaneToggle('enableClassicContext', {
                  label: 'Enable classic _spPageContextInfo', onText: 'Enabled', offText: 'Disabled'
                }),
                PropertyPaneToggle('useIframe', {
                  label: 'Run in sandboxed iFrame (recommended)', onText: 'iFrame', offText: 'Inline'
                }),
                PropertyPaneToggle('domSanitize', {
                  label: 'Sanitize HTML before render', onText: 'On', offText: 'Off'
                }),
                PropertyPaneToggle('allowInlineScript', {
                  label: 'Allow inline <script> execution', onText: 'Allowed', offText: 'Blocked'
                }),
                PropertyPaneToggle('cspStrict', {
                  label: 'Strict CSP (blocks unsafe-inline in iFrame)', onText: 'Strict', offText: 'Relaxed'
                }),
                PropertyPaneTextField('allowedDomains', {
                  label: 'Allowed script domains (comma-separated)',
                  description: 'Examples: self, *.sharepoint.com, cdn.contoso.com'
                }),
                codeField
              ]
            }
          ]
        }
      ]
    };
  }
}
