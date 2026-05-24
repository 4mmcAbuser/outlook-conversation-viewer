# Outlook Conversation History JSON Viewer Add-in

A premium Outlook read-mode taskpane add-in that retrieves the complete history of an email thread conversation and renders it in a beautiful, collapsible, syntax-highlighted JSON tree and timeline view.

![Glassmorphic Design Mockup](https://raw.githubusercontent.com/github/outlook-conversation-viewer/main/assets/mockup.png)

## Core Architectural Design

- **Exchange Web Services (EWS) SOAP Engine**: Rather than relying on Microsoft Graph API which requires Azure Active Directory App Registration, global admin consent, and multi-tenant OAuth configurations, this add-in utilizes direct EWS SOAP calls (`GetConversationItems` operation) via `makeEwsRequestAsync`. 
- **Sandboxed Security**: It operates entirely in the Outlook client sandbox using the current user's active session, allowing for frictionless integration in large enterprise environments.
- **Mock Developer Mode**: When opened outside the Outlook container (e.g., standard browser window during development), the script automatically boots into a feature-rich, interactive **Preview Mode** using mock data so you can test and preview the styling, layout, timeline rendering, and JSON tree behaviors without requiring Exchange Server hosting.

## Project Structure

```
outlook-conversation-viewer/
├── manifest.xml       # Outlook Add-in XML configuration
├── taskpane.html      # Taskpane layout with widgets, tabs, timeline, and JSON panel
├── taskpane.css       # Stunning HSL-tailored glassmorphic dark-theme styles
├── taskpane.js        # Core controller (EWS querying, XML parsing, JSON structuring, Copy/Download actions)
└── README.md          # Setup & Deployment Instructions
```

## Styling Features (Rich Aesthetics)

- **Palette**: Tailored dark-mode using deep harmonious carbon HSL variables (`hsl(222, 47%, 9%)`).
- **Glassmorphism**: Backdrop blur filter with semi-transparent light borders (`rgba(255, 255, 255, 0.08)`) and high depth shadows.
- **timeline Timeline View**: Chronological ordering of events with glowing active nodes.
- **Collapsible JSON Highlighter**: Fully dynamic object browser with collapsible bracket groups, color-highlighted keys, and variables.
- **Micro-Animations**: Transitions on button triggers (like Copy changing to Connected green with scale changes) and dynamic loading bars.

## How to Run & Sideload in Outlook

### Step 1: Run Locally
To load the files in Outlook, they must be served over a secure connection (`https://`).
You can spin up a local HTTPS server inside this folder:
```bash
# Using local-ssl-proxy or http-server with local certs
npx http-server -S -C cert.pem -K key.pem -p 3000
```
*Note: If testing without an actual Outlook environment, simply double-click or open `taskpane.html` directly in a browser to trigger the built-in **Preview Mode**.*

### Step 2: Sideload in Outlook on the Web (OWA)
1. Open [Outlook on the Web](https://outlook.office.com/).
2. Go to **Settings (Gear Icon)** -> **General** -> **Manage add-ins**.
3. Select **My add-ins** tab.
4. Under **Custom add-ins**, click **Add a custom add-in** -> **Add from file...**
5. Upload the [manifest.xml](file:///C:/Users/user/.gemini/antigravity/scratch/outlook-conversation-viewer/manifest.xml) file.
6. Open any email, click the three dots (`...`) in the action ribbon on the top-right of the email, and click **View Conversation JSON** to slide out the gorgeous insights dashboard!

---

## License
MIT License. Created by Premium Add-ins Corp.
