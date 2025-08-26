# ScriptEditor (SPFx)
Modern Script Editor web part for SharePoint Online.
Script Editor SPFx Web Part – Technical, Architecture & Development Guide
Scope: Modern Script Editor–style web part implemented on SPFx 1.19.0 with hardened security and PnP property controls
________________________________________
1. Summary
This project delivers a modern Script Editor web part for SharePoint Online that allows controlled rendering of HTML/CSS/JavaScript in site pages. The design replicates the UX of the community Modern Script Editor while adding explicit security hardening:
•	Two render modes: sandboxed iFrame (recommended) and Inline.
•	Sanitization with DOMPurify prior to render.
•	Allow list for external script sources.
•	Content Security Policy (CSP) when using iFrame.
•	Optional exposure of classic _spPageContextInfo for legacy scripts.
•	Built in code editor (PnP PropertyFieldCodeEditor) for a better authoring experience.
Why iFrame by default? It isolates custom markup from the page DOM, and our CSP + sandbox attributes reduce cross scope impact even if the markup is malicious or buggy.
________________________________________
2. Environment & Versions
•	SPFx: 1.19.0
•	Node.js: 18.19.1
•	npm: 10.2.4
•	Gulp: CLI 2.3.0 / Local 4.0.2
•	TypeScript: 5.4.5
•	PnP Property Controls: @pnp/spfx-property-controls@^3.20.0
•	UI Fabric/Fluent: @fluentui/react@^8 (peer for property controls)
•	Sanitizer: dompurify@^3
PnP core libraries (@pnp/sp, etc.) are not required for this web part.
________________________________________
3. High Level Architecture
3.1 Components
•	Web part: ScriptEditorWebPart (TypeScript).
•	Property Pane: PnP Code Editor + toggles for security/behavior.
•	Renderer: chooses between Inline and iFrame strategies.
•	Security utilities: allow list validator, CSP assembler, sanitizer adapter.
3.2 Data Flow
1.	Author configures the web part in property pane (markup, iFrame/Inline, allow list, etc.).
2.	On render, markup is optionally sanitized → rendered inside iFrame (with CSP) or inline.
3.	In Inline mode:
a. Extract <script src="…"> → validate each URL by allow list → SPComponentLoader.loadScript.
b. If enabled, inject inline <script> blocks after outer HTML is set.
3.3 Key Files
•	src/webparts/scriptEditor/ScriptEditorWebPart.ts – main implementation.
•	config/config.json – bundle definition (entrypoint + manifest).
•	src/webparts/scriptEditor/ScriptEditorWebPart.manifest.json – web part identity & defaults.
•	config/tsconfig.json & /tsconfig.json – TypeScript targets (ES2017, DOM).
________________________________________
4. Property Model (Authoring Experience)
Property	Type	Default	Purpose
editTitle	string	(empty)	Small title shown only in Edit mode above the part.
keepPadding	boolean	true	Keep or remove top/bottom padding around the host control.
enableClassicContext	boolean	false	Expose _spPageContextInfo (legacy compatibility).
markup	string	(sample DIV)	HTML/JS/CSS content edited via the { } panel.
useIframe	boolean	true	Render in sandboxed iFrame (recommended).
domSanitize	boolean	true	Run DOMPurify sanitization prior to render.
allowInlineScript	boolean	false	Allow inline <script> (blocked when cspStrict is on).
cspStrict	boolean	true	In iFrame, omit 'unsafe-inline' from script-src.
allowedDomains	string	self, *.sharepoint.com	Comma separated allow list for external <script src="…"> URLs.
The property pane shows a placeholder tile saying “Please configure the web part” with an Edit markup button if markup is empty.
________________________________________
5. Rendering Strategies
5.1 iFrame Mode (Default)
•	Sandbox: allow-scripts allow-same-origin to enable JS but keep isolation.
•	CSP meta tag injected into the iFrame document, for example:
•	default-src 'self';
•	script-src 'self' cdn.contoso.com  ( + 'unsafe-inline' when strict=false & allowInlineScript=true )
•	style-src 'self' 'unsafe-inline' data:;
•	img-src   'self' data:;
•	connect-src 'self' cdn.contoso.com
•	Sanitization: If domSanitize=true, HTML is run through DOMPurify before writing to the iFrame.
•	Auto resize: Reads body/document scroll heights and adjusts the iFrame height.
5.2 Inline Mode
•	Parses markup with DOMParser.
•	Extracts all <script> elements:
o	For each src: validate via allow list, then SPComponentLoader.loadScript.
o	For inline code: inject only if allowInlineScript=true.
•	Sanitizes outer HTML (forbids <script> during sanitize pass) when domSanitize=true.
________________________________________
6. Security Hardening
6.1 Threat Model
•	XSS via untrusted markup or scripts.
•	Lateral impact on the hosting page (CSS/DOM pollution, global overrides).
•	External script abuse from unapproved domains.
6.2 Controls
1.	Isolated Execution (iFrame) – primary containment boundary.
2.	CSP – explicit script-src & friends inside the iFrame.
3.	Sanitization – DOMPurify removes dangerous markup before write.
4.	Allow list Enforcement – only approved external script hosts load.
5.	Inline Script Toggle – disabled by default and blocked under strict CSP.
6.	Legacy Context Toggle – _spPageContextInfo is opt in and intended for legacy code only.
6.3 Governance Recommendations
•	Restrict who can author pages using this web part (Owners/Designers only).
•	Maintain a central allow list of corporate CDNs (e.g., *.contoso.com).
•	Prefer iFrame + strict CSP for production pages.
•	Review and version custom scripts; avoid referencing 3rd party libraries that auto update.
________________________________________
7. Troubleshooting
Symptom: “Something went wrong – [object Object]”
•	Open DevTools → Console, expand the error to view .message/.stack.
•	Verify config/config.json has correct entrypoint & manifest paths.
•	Confirm temp/manifests.js contains your web part id and the entryModuleId matches the bundle name.
•	Run gulp clean && gulp build, and reload workbench with SW disabled.
•	Ensure dependencies: @pnp/spfx-property-controls, @fluentui/react, dompurify.
Symptom: {} button missing in property pane
•	We use static import of PnP Code Editor; if still missing, clear caches and verify dependency version ≥ 3.20.0.
Existing instance still shows old sample text
•	Page stored properties override manifest defaults. Clear the Code field and Apply, or remove/re add the part.
________________________________________
8. Risks & Limitations
•	Custom code risk: The web part enforces controls, but unsafe author content can still misbehave (especially in Inline mode).
•	CSP edge cases: Some libraries require unsafe-inline. Use Inline mode sparingly or engineer a safer initialization pattern.
•	Performance: Multiple heavy external scripts can impact page load; prefer bundling or approved CDNs.
________________________________________
 Appendix
1 Property Pane – Quick Reference
•	Run in sandboxed iFrame: On (recommended)
•	Sanitize HTML before render: On
•	Allow inline <script>: Off
•	Strict CSP: On
•	Allowed domains: self, *.sharepoint.com, cdn.contoso.com
2 Example Allow list Policies
•	Strict (internal only): self, *.sharepoint.com
•	Internal CDN: self, *.sharepoint.com, cdn.contoso.com
•	Add Azure Static Web Apps: self, *.sharepoint.com, *.azureedge.net (review ownership first)
3 Minimal Authoring Examples
A. Simple HTML
<div class="ms-Grid-row" style="padding:8px;border:1px dashed #ccc">Hello from Script Editor</div>
B. Approved external script
<div id="hello"></div>
<script src="https://cdn.contoso.com/libs/hello.min.js"></script>
<script>
  Hello.render('#hello');
</script>
4 Key Code Decisions
•	Static import of PropertyFieldCodeEditor avoids dynamic import CSP issues.
•	DOMPurify adapter supports both default and namespace exports to prevent runtime import shape errors.
•	iFrame CSP assembled per settings; sandbox enforces isolation boundary.
•	Inline mode uses SPComponentLoader for predictable, SPFx aware script loading.
________________________________________
Script Editor pane Settings:
•	Enable classic _spPageContextInfo → On
•	Run in sandboxed iFrame → Off (turn this off for easiest context/auth)
•	Sanitize HTML before render → Off
•	Allow inline <script> execution → Allowed
•	Strict CSP → Permissive
•	Allowed script domains → self, *.sharepoint.com
Click Apply, then paste the code.
If you must keep “iFrame = On”, this code still tries window.top._spPageContextInfo first.
________________________________________
Screenshots:


<img width="342" height="847" alt="image" src="https://github.com/user-attachments/assets/61eb216c-203d-42cd-b029-7166ab397594" />
