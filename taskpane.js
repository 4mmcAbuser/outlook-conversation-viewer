// Global Variables
let rawConversationData = null;
let rawJsonResponse = "";
let isOfficeInitialized = false;

// Mock Data for Browser Preview (Premium Developer Mode)
const mockConversation = {
    conversationId: "CONV_ID_MOCK_8d2f1f50",
    subject: "Re: Design Proposal - Premium Glassmorphism Theme",
    totalMessages: 3,
    participants: ["Sarah Jenkins", "Alex Rivera", "Emily Chen"],
    timeframe: "May 24, 2026",
    messages: [
        {
            Subject: "Design Proposal - Premium Glassmorphism Theme",
            From: { EmailAddress: { Name: "Sarah Jenkins", Address: "sarah.j@company.com" } },
            ToRecipients: [{ EmailAddress: { Name: "Team", Address: "team@company.com" } }],
            ReceivedDateTime: "2026-05-24T14:20:00Z",
            BodyPreview: "Hi team, I've put together the visual concept for our new client interface. We are aiming for a rich glassmorphism aesthetic with tailored HSL variables, backdrop blur, and semi-transparent borders.",
            HasAttachments: true
        },
        {
            Subject: "Re: Design Proposal - Premium Glassmorphism Theme",
            From: { EmailAddress: { Name: "Alex Rivera", Address: "alex.r@company.com" } },
            ToRecipients: [{ EmailAddress: { Name: "Sarah Jenkins", Address: "sarah.j@company.com" } }],
            ReceivedDateTime: "2026-05-24T15:05:00Z",
            BodyPreview: "This looks stunning, Sarah! The micro-animations on hover and the dark mode gradients feel extremely premium. Performance-wise, we just need to ensure the backdrop-filter styles don't bottleneck mobile rendering.",
            HasAttachments: false
        },
        {
            Subject: "Re: Design Proposal - Premium Glassmorphism Theme",
            From: { EmailAddress: { Name: "Emily Chen", Address: "emily.c@company.com" } },
            ToRecipients: [{ EmailAddress: { Name: "Sarah Jenkins", Address: "sarah.j@company.com" } }, { EmailAddress: { Name: "Alex Rivera", Address: "alex.r@company.com" } }],
            ReceivedDateTime: "2026-05-24T15:30:00Z",
            BodyPreview: "Incredible work, both of you. The client is going to be absolutely wowed. Let's lock in this design direction. Alex, let me know how the sandbox test goes so we can schedule the dev sprint.",
            HasAttachments: false
        }
    ]
};

// Office Init
Office.onReady((info) => {
    isOfficeInitialized = true;
    const statusBadge = document.getElementById("app-status");
    
    if (info.host === Office.HostType.Outlook) {
        statusBadge.textContent = "Outlook Host";
        statusBadge.className = "status-badge connected";
        initializeAddIn();
    } else {
        statusBadge.textContent = "Preview Mode";
        statusBadge.className = "status-badge connected";
        loadMockData();
    }
});

// Setup Event Listeners
document.addEventListener("DOMContentLoaded", () => {
    // Tab switching
    const tabBtns = document.querySelectorAll(".tab-btn");
    tabBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            tabBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            
            const targetTab = btn.getAttribute("data-tab");
            const tabContents = document.querySelectorAll(".tab-content");
            tabContents.forEach(content => content.classList.remove("active"));
            document.getElementById(`tab-${targetTab}`).classList.add("active");
            
            if (targetTab === "source-code") {
                loadSourceCode();
            }
        });
    });

    // Action buttons
    document.getElementById("btn-copy-json").addEventListener("click", copyJsonToClipboard);
    document.getElementById("btn-download-json").addEventListener("click", downloadJsonFile);
    document.getElementById("btn-copy-xml").addEventListener("click", copyRawToClipboard);
    document.getElementById("btn-collapse-all").addEventListener("click", () => renderJsonView(rawConversationData, false));
    document.getElementById("btn-expand-all").addEventListener("click", () => renderJsonView(rawConversationData, true));
    
    // Source code viewers
    document.getElementById("source-file-selector").addEventListener("change", loadSourceCode);
    document.getElementById("btn-copy-source").addEventListener("click", copySourceCodeToClipboard);
});

// Primary Add-in Entry Point inside Outlook
function initializeAddIn() {
    const loader = document.getElementById("loader");
    const errorPanel = document.getElementById("error-panel");
    const dashboard = document.getElementById("dashboard");
    
    loader.classList.remove("hidden");
    errorPanel.classList.add("hidden");
    dashboard.classList.add("hidden");

    try {
        const item = Office.context.mailbox.item;
        if (!item) {
            showError("No email conversation loaded.");
            return;
        }

        const conversationId = item.conversationId;
        if (!conversationId) {
            showError("Could not retrieve conversation ID for this message.");
            return;
        }

        // Try EWS SOAP first (corrected schema), fall back to REST if it fails
        fetchConversationViaEWS(conversationId);
    } catch (e) {
        showError(`Initialization error: ${e.message}`);
    }
}

// ============================================================
// REST API Engine (replaces legacy EWS SOAP)
// Works on ALL accounts including free Outlook.com
// ============================================================
function fetchConversationViaREST() {
    const progressBar = document.getElementById("progress-bar");
    progressBar.style.width = "30%";

    // Step 1: Get a REST callback token
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        progressBar.style.width = "50%";

        if (result.status !== "succeeded") {
            showError("Authentication failed: " + (result.error ? result.error.message : "Could not obtain REST token."));
            return;
        }

        const accessToken = result.value;
        const restUrl = Office.context.mailbox.restUrl;
        const conversationId = Office.context.mailbox.item.conversationId;

        progressBar.style.width = "70%";

        // Step 2: Query the REST API for all messages in this conversation
        const apiUrl = `${restUrl}/api/v2.0/me/messages?$filter=ConversationId eq '${conversationId}'&$orderby=ReceivedDateTime asc&$select=Subject,From,ToRecipients,ReceivedDateTime,BodyPreview,HasAttachments&$top=50`;

        fetch(apiUrl, {
            method: "GET",
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json"
            }
        })
        .then(response => {
            progressBar.style.width = "90%";
            if (!response.ok) {
                return response.text().then(text => {
                    throw new Error(`HTTP ${response.status}: ${text.substring(0, 200)}`);
                });
            }
            return response.json();
        })
        .then(data => {
            progressBar.style.width = "100%";
            const messages = data.value || [];
            
            if (messages.length === 0) {
                showError("No messages found in this conversation thread.");
                return;
            }

            // Store the raw JSON response
            rawJsonResponse = JSON.stringify(data, null, 2);

            // Build the structured conversation object
            const participantsSet = new Set();
            messages.forEach(msg => {
                if (msg.From && msg.From.EmailAddress) {
                    participantsSet.add(msg.From.EmailAddress.Name || msg.From.EmailAddress.Address);
                }
                if (msg.ToRecipients) {
                    msg.ToRecipients.forEach(r => {
                        if (r.EmailAddress) {
                            participantsSet.add(r.EmailAddress.Name || r.EmailAddress.Address);
                        }
                    });
                }
            });

            rawConversationData = {
                conversationId: conversationId,
                subject: Office.context.mailbox.item.subject || messages[0].Subject || "No Subject",
                totalMessages: messages.length,
                participants: Array.from(participantsSet),
                timeframe: getFormattedTimeframe(messages.map(m => m.ReceivedDateTime)),
                messages: messages
            };

            renderDashboard(rawConversationData, rawJsonResponse);
        })
        .catch(error => {
            showError("REST API Error: " + error.message);
        });
    });
}

// Helper to format conversation timeframe
function getFormattedTimeframe(dates) {
    if (!dates || dates.length === 0) return "";
    const validDates = dates.filter(d => d);
    if (validDates.length === 0) return "";
    if (validDates.length === 1) return formatDate(validDates[0]);
    return `${formatDate(validDates[0])} - ${formatDate(validDates[validDates.length - 1])}`;
}

function formatDate(dateString) {
    try {
        const d = new Date(dateString);
        return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
    } catch {
        return "";
    }
}

// Show details on Dashboard
function renderDashboard(data, jsonResponse) {
    document.getElementById("loader").classList.add("hidden");
    document.getElementById("error-panel").classList.add("hidden");
    
    const dashboard = document.getElementById("dashboard");
    dashboard.classList.remove("hidden");

    // Populate widgets
    document.getElementById("conv-subject").textContent = data.subject;
    document.getElementById("stat-emails").textContent = data.totalMessages;
    document.getElementById("stat-people").textContent = data.participants.length || 1;

    // Render Timeline View
    renderTimelineView(data.messages);

    // Render JSON Tree View
    renderJsonView(data, true);

    // Populate Raw JSON response view
    document.getElementById("xml-output").textContent = jsonResponse || "No raw response available.";
}

// Renders the visual Timeline tab
function renderTimelineView(messages) {
    const timelineList = document.getElementById("timeline-list");
    timelineList.innerHTML = "";

    messages.forEach((msg, index) => {
        const item = document.createElement("li");
        item.className = "timeline-item";
        item.style.animationDelay = `${index * 0.1}s`;

        const dateStr = msg.ReceivedDateTime ? (() => {
            const d = new Date(msg.ReceivedDateTime);
            return d.toLocaleDateString("en-US", { month: "short", day: "numeric" }) + " " + 
                   d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
        })() : "Unknown date";

        const senderName = msg.From && msg.From.EmailAddress 
            ? (msg.From.EmailAddress.Name || msg.From.EmailAddress.Address) 
            : "Unknown";

        const attachmentBadge = msg.HasAttachments ? ' 📎' : '';

        item.innerHTML = `
            <div class="timeline-marker"></div>
            <div class="timeline-content">
                <div class="timeline-header">
                    <span class="sender-name">${escapeHtml(senderName)}${attachmentBadge}</span>
                    <span class="date-text">${dateStr}</span>
                </div>
                <div class="snippet-text">${escapeHtml(msg.BodyPreview || "No content preview available.")}</div>
            </div>
        `;
        timelineList.appendChild(item);
    });
}

// Interactive Collapsible JSON Renderer
function renderJsonView(obj, defaultExpanded = true) {
    const jsonOutput = document.getElementById("json-output");
    jsonOutput.innerHTML = "";
    
    const html = formatJsonHtml(obj, defaultExpanded, 0);
    jsonOutput.appendChild(html);
}

function formatJsonHtml(val, expanded, depth) {
    const container = document.createElement("div");
    
    if (val === null || val === undefined) {
        const span = document.createElement("span");
        span.className = "json-null";
        span.textContent = "null";
        container.appendChild(span);
    } else if (typeof val === "boolean") {
        const span = document.createElement("span");
        span.className = "json-boolean";
        span.textContent = val.toString();
        container.appendChild(span);
    } else if (typeof val === "number") {
        const span = document.createElement("span");
        span.className = "json-number";
        span.textContent = val.toString();
        container.appendChild(span);
    } else if (typeof val === "string") {
        const span = document.createElement("span");
        span.className = "json-string";
        span.textContent = `"${escapeHtml(val)}"`;
        container.appendChild(span);
    } else if (Array.isArray(val)) {
        if (val.length === 0) {
            container.appendChild(document.createTextNode("[]"));
            return container;
        }

        const shouldExpand = expanded && depth < 3;
        const block = document.createElement("span");
        const openBracket = document.createElement("span");
        openBracket.className = "collapsible-trigger";
        openBracket.textContent = shouldExpand ? "[" : `[ ...${val.length} items ]`;
        
        const closeBracket = document.createElement("span");
        closeBracket.textContent = "]";
        
        const contents = document.createElement("div");
        contents.className = "json-node";
        if (!shouldExpand) contents.style.display = "none";

        openBracket.onclick = () => {
            if (contents.style.display === "none") {
                contents.style.display = "block";
                openBracket.textContent = "[";
            } else {
                contents.style.display = "none";
                openBracket.textContent = `[ ...${val.length} items ]`;
            }
        };

        val.forEach((item, index) => {
            const itemDiv = document.createElement("div");
            itemDiv.appendChild(formatJsonHtml(item, expanded, depth + 1));
            if (index < val.length - 1) {
                itemDiv.appendChild(document.createTextNode(","));
            }
            contents.appendChild(itemDiv);
        });

        block.appendChild(openBracket);
        block.appendChild(contents);
        block.appendChild(closeBracket);
        container.appendChild(block);
    } else if (typeof val === "object") {
        const keys = Object.keys(val);
        if (keys.length === 0) {
            container.appendChild(document.createTextNode("{}"));
            return container;
        }

        const shouldExpand = expanded && depth < 3;
        const block = document.createElement("span");
        const openBrace = document.createElement("span");
        openBrace.className = "collapsible-trigger";
        openBrace.textContent = shouldExpand ? "{" : "{ ... }";
        
        const closeBrace = document.createElement("span");
        closeBrace.textContent = "}";
        
        const contents = document.createElement("div");
        contents.className = "json-node";
        if (!shouldExpand) contents.style.display = "none";

        openBrace.onclick = () => {
            if (contents.style.display === "none") {
                contents.style.display = "block";
                openBrace.textContent = "{";
            } else {
                contents.style.display = "none";
                openBrace.textContent = "{ ... }";
            }
        };

        keys.forEach((key, index) => {
            const rowDiv = document.createElement("div");
            const keySpan = document.createElement("span");
            keySpan.className = "json-key";
            keySpan.textContent = `"${key}"`;
            
            rowDiv.appendChild(keySpan);
            rowDiv.appendChild(document.createTextNode(": "));
            rowDiv.appendChild(formatJsonHtml(val[key], expanded, depth + 1));
            if (index < keys.length - 1) {
                rowDiv.appendChild(document.createTextNode(","));
            }
            contents.appendChild(rowDiv);
        });

        block.appendChild(openBrace);
        block.appendChild(contents);
        block.appendChild(closeBrace);
        container.appendChild(block);
    }

    return container;
}

// Copy JSON Actions with micro-animations
function copyJsonToClipboard() {
    if (!rawConversationData) return;
    const jsonStr = JSON.stringify(rawConversationData, null, 2);
    
    navigator.clipboard.writeText(jsonStr).then(() => {
        const btn = document.getElementById("btn-copy-json");
        const textSpan = btn.querySelector(".btn-text");
        
        textSpan.textContent = "Copied ✓";
        btn.style.background = "linear-gradient(135deg, #10b981 0%, #059669 100%)";
        btn.style.boxShadow = "0 4px 12px rgba(16, 185, 129, 0.3)";
        
        setTimeout(() => {
            textSpan.textContent = "Copy JSON";
            btn.style.background = "";
            btn.style.boxShadow = "";
        }, 2200);
    });
}

function copyRawToClipboard() {
    if (!rawJsonResponse) return;
    navigator.clipboard.writeText(rawJsonResponse).then(() => {
        const btn = document.getElementById("btn-copy-xml");
        btn.textContent = "Copied ✓";
        setTimeout(() => {
            btn.textContent = "Copy Raw";
        }, 2200);
    });
}

// Download local JSON file
function downloadJsonFile() {
    if (!rawConversationData) return;
    const jsonStr = JSON.stringify(rawConversationData, null, 2);
    const blob = new Blob([jsonStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement("a");
    a.href = url;
    a.download = `conversation-${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Show errors
function showError(message) {
    document.getElementById("loader").classList.add("hidden");
    document.getElementById("dashboard").classList.add("hidden");
    
    const errorPanel = document.getElementById("error-panel");
    errorPanel.classList.remove("hidden");
    
    document.getElementById("error-message").textContent = message;
    
    const statusBadge = document.getElementById("app-status");
    statusBadge.textContent = "Error";
    statusBadge.className = "status-badge error";
}

// Html escape helper
function escapeHtml(str) {
    if (!str) return "";
    return str.toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// Load Mock Data (Preview Mode outside Outlook)
function loadMockData() {
    rawConversationData = {
        conversationId: mockConversation.conversationId,
        subject: mockConversation.subject,
        totalMessages: mockConversation.totalMessages,
        participants: mockConversation.participants,
        timeframe: mockConversation.timeframe,
        messages: mockConversation.messages
    };
    rawJsonResponse = JSON.stringify(mockConversation, null, 2);
    renderDashboard(rawConversationData, rawJsonResponse);
}

// Load and display source code
function loadSourceCode() {
    const file = document.getElementById("source-file-selector").value;
    const output = document.getElementById("source-output");
    output.textContent = "Loading " + file + "...";
    
    fetch("./" + file)
        .then(response => {
            if (!response.ok) throw new Error("Could not fetch file: " + response.statusText);
            return response.text();
        })
        .then(text => {
            output.textContent = text;
        })
        .catch(err => {
            output.textContent = "Error loading " + file + ": " + err.message;
        });
}

function copySourceCodeToClipboard() {
    const output = document.getElementById("source-output");
    const code = output.textContent;
    if (code.startsWith("Loading ") || code.startsWith("Error loading ")) return;
    
    navigator.clipboard.writeText(code).then(() => {
        const btn = document.getElementById("btn-copy-source");
        btn.textContent = "Copied ✓";
        setTimeout(() => {
            btn.textContent = "Copy Code";
        }, 2200);
    });
}

// ============================================================
// EWS SOAP Engine (corrected schema, works universally)
// ============================================================
function fetchConversationViaEWS(conversationId) {
    const progressBar = document.getElementById("progress-bar");
    progressBar.style.width = "25%";

    const soapRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:GetConversationItems>
      <m:ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="message:Sender" />
          <t:FieldURI FieldURI="item:Body" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemRequests>
        <t:ConversationRequest>
          <t:ConversationId Id="${conversationId}" />
        </t:ConversationRequest>
      </m:ItemRequests>
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;

    progressBar.style.width = "45%";

    try {
        Office.context.mailbox.makeEwsRequestAsync(soapRequest, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                progressBar.style.width = "85%";
                parseEwsSoapResponse(asyncResult.value);
            } else {
                console.warn("EWS SOAP Call failed: " + (asyncResult.error ? asyncResult.error.message : "unknown") + ". Falling back to REST API.");
                fetchConversationViaREST();
            }
        });
    } catch (err) {
        console.warn("EWS invocation threw error: " + err.message + ". Falling back to REST API.");
        fetchConversationViaREST();
    }
}

function parseEwsSoapResponse(xmlString) {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlString, "text/xml");
        
        // Check for SOAP faults
        const faultNode = xmlDoc.getElementsByTagNameNS("http://schemas.xmlsoap.org/soap/envelope/", "Fault")[0] ||
                          xmlDoc.getElementsByTagName("soap:Fault")[0] ||
                          xmlDoc.getElementsByTagName("Fault")[0];
        if (faultNode) {
            const faultString = faultNode.getElementsByTagName("faultstring")[0]?.textContent || "EWS SOAP Fault occurred.";
            console.warn("EWS Fault: " + faultString + ". Falling back to REST API.");
            fetchConversationViaREST();
            return;
        }

        // Get conversation nodes
        const conversationNodes = xmlDoc.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "ConversationNode") ||
                                  xmlDoc.getElementsByTagName("t:ConversationNode") ||
                                  xmlDoc.getElementsByTagName("ConversationNode");

        if (conversationNodes.length === 0) {
            console.warn("No conversation history node structure found in EWS response. Falling back to REST API.");
            fetchConversationViaREST();
            return;
        }

        const messages = [];
        const participantsSet = new Set();
        let mainSubject = "";

        // Iterate through EWS nodes
        for (let i = 0; i < conversationNodes.length; i++) {
            const node = conversationNodes[i];
            const itemNodes = node.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "Message") ||
                              node.getElementsByTagName("t:Message") ||
                              node.getElementsByTagName("Message");

            for (let j = 0; j < itemNodes.length; j++) {
                const item = itemNodes[j];
                
                const itemId = getXmlNodeValue(item, "ItemId", "Id");
                const subject = getXmlNodeText(item, "Subject");
                const dateTimeReceived = getXmlNodeText(item, "DateTimeReceived");
                
                // Get Sender Info
                let senderName = "Unknown";
                let senderEmail = "unknown@domain.com";
                const senderNode = item.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "Sender")[0] ||
                                   item.getElementsByTagName("t:Sender")[0] ||
                                   item.getElementsByTagName("Sender")[0];
                if (senderNode) {
                    senderName = getXmlNodeText(senderNode, "Name");
                    senderEmail = getXmlNodeText(senderNode, "EmailAddress");
                }
                
                if (senderName && senderName !== "Unknown") {
                    participantsSet.add(senderName);
                }

                const hasAttachments = getXmlNodeText(item, "HasAttachments") === "true";

                // Get body content
                let bodySnippet = getXmlNodeText(item, "Body");
                if (bodySnippet && bodySnippet.length > 500) {
                    bodySnippet = bodySnippet.substring(0, 500) + "...";
                }

                if (!mainSubject && subject) {
                    mainSubject = subject;
                }

                // Map to REST API compatibility structure
                messages.push({
                    Subject: subject || "No Subject",
                    From: {
                        EmailAddress: {
                            Name: senderName,
                            Address: senderEmail
                        }
                    },
                    ToRecipients: [],
                    ReceivedDateTime: dateTimeReceived || new Date().toISOString(),
                    BodyPreview: bodySnippet || "",
                    HasAttachments: hasAttachments
                });
            }
        }

        // Sort chronological
        messages.sort((a, b) => new Date(a.ReceivedDateTime) - new Date(b.ReceivedDateTime));

        rawConversationData = {
            conversationId: Office.context.mailbox.item.conversationId,
            subject: mainSubject || Office.context.mailbox.item.subject || "Subject Unavailable",
            totalMessages: messages.length,
            participants: Array.from(participantsSet),
            timeframe: getFormattedTimeframe(messages.map(m => m.ReceivedDateTime)),
            messages: messages
        };

        // Formats the XML nicely for display in the tab
        rawJsonResponse = JSON.stringify(rawConversationData, null, 2);

        const progressBar = document.getElementById("progress-bar");
        progressBar.style.width = "100%";

        renderDashboard(rawConversationData, xmlString);
    } catch (e) {
        console.warn("Parsing XML response failed: " + e.message + ". Falling back to REST API.");
        fetchConversationViaREST();
    }
}

// Utility to get XML text content handling namespaces
function getXmlNodeText(parentNode, localName) {
    const nodes = parentNode.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", localName);
    if (nodes.length > 0) return nodes[0].textContent;
    
    const prefNodes = parentNode.getElementsByTagName("t:" + localName);
    if (prefNodes.length > 0) return prefNodes[0].textContent;
    
    const plainNodes = parentNode.getElementsByTagName(localName);
    if (plainNodes.length > 0) return plainNodes[0].textContent;
    
    return "";
}

// Utility to get XML attributes handling namespaces
function getXmlNodeValue(parentNode, localName, attributeName) {
    const nodes = parentNode.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", localName);
    if (nodes.length > 0) return nodes[0].getAttribute(attributeName);
    
    const prefNodes = parentNode.getElementsByTagName("t:" + localName);
    if (prefNodes.length > 0) return prefNodes[0].getAttribute(attributeName);
    
    const plainNodes = parentNode.getElementsByTagName(localName);
    if (plainNodes.length > 0) return plainNodes[0].getAttribute(attributeName);
    
    return "";
}
