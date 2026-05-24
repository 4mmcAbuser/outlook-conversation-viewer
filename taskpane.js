// Global Variables
let rawConversationData = null;
let rawXmlResponse = "";
let isOfficeInitialized = false;

// Mock Data for Browser Preview (Premium Developer Mode)
const mockConversation = {
    conversationId: "CONV_ID_MOCK_8d2f1f50-32df-4277-bc60-0a370e5b565d",
    subject: "Re: Design Proposal - Premium Glassmorphism Theme",
    totalMessages: 3,
    participants: ["Sarah Jenkins (Design Lead)", "Alex Rivera (Tech Lead)", "Emily Chen (Product Manager)"],
    timeframe: "May 24, 2026",
    messages: [
        {
            messageId: "MSG_ID_1",
            subject: "Design Proposal - Premium Glassmorphism Theme",
            sender: {
                name: "Sarah Jenkins",
                email: "sarah.j@company.com"
            },
            dateTimeReceived: "2026-05-24T14:20:00Z",
            bodySnippet: "Hi team, I've put together the visual concept for our new client interface. We are aiming for a rich glassmorphism aesthetic with tailored HSL variables, backdrop blur, and semi-transparent borders. Please review the attached mockup assets."
        },
        {
            messageId: "MSG_ID_2",
            subject: "Re: Design Proposal - Premium Glassmorphism Theme",
            sender: {
                name: "Alex Rivera",
                email: "alex.r@company.com"
            },
            dateTimeReceived: "2026-05-24T15:05:00Z",
            bodySnippet: "This looks stunning, Sarah! The micro-animations on hover and the dark mode gradients feel extremely premium. Performance-wise, we just need to ensure the backdrop-filter styles don't bottleneck mobile rendering. I'll create a sandbox mockup to test."
        },
        {
            messageId: "MSG_ID_3",
            subject: "Re: Design Proposal - Premium Glassmorphism Theme",
            sender: {
                name: "Emily Chen",
                email: "emily.c@company.com"
            },
            dateTimeReceived: "2026-05-24T15:30:00Z",
            bodySnippet: "Incredible work, both of you. The client is going to be absolutely wowed. Let's lock in this design direction. Alex, let me know how the sandbox test goes so we can schedule the dev sprint."
        }
    ]
};

const mockXmlResponse = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <m:GetConversationItemsResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseMessages>
        <m:GetConversationItemsResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:Conversation>
            <t:ConversationId Id="CONV_ID_MOCK_8d2f1f50-32df-4277-bc60-0a370e5b565d"/>
            <t:ConversationNodes>
              <t:ConversationNode>
                <t:InternetMessageId>&lt;sarah.j@company.com&gt;</t:InternetMessageId>
                <t:Items>
                  <t:Message>
                    <t:ItemId Id="MSG_ID_1"/>
                    <t:Subject>Design Proposal - Premium Glassmorphism Theme</t:Subject>
                    <t:DateTimeReceived>2026-05-24T14:20:00Z</t:DateTimeReceived>
                    <t:Sender>
                      <t:Mailbox>
                        <t:Name>Sarah Jenkins</t:Name>
                        <t:EmailAddress>sarah.j@company.com</t:EmailAddress>
                      </t:Mailbox>
                    </t:Sender>
                    <t:Body BodyType="Text">Hi team, I've put together the visual concept for our new client interface. We are aiming for a rich glassmorphism aesthetic with tailored HSL variables, backdrop blur, and semi-transparent borders. Please review the attached mockup assets.</t:Body>
                  </t:Message>
                </t:Items>
              </t:ConversationNode>
              <t:ConversationNode>
                <t:InternetMessageId>&lt;alex.r@company.com&gt;</t:InternetMessageId>
                <t:Items>
                  <t:Message>
                    <t:ItemId Id="MSG_ID_2"/>
                    <t:Subject>Re: Design Proposal - Premium Glassmorphism Theme</t:Subject>
                    <t:DateTimeReceived>2026-05-24T15:05:00Z</t:DateTimeReceived>
                    <t:Sender>
                      <t:Mailbox>
                        <t:Name>Alex Rivera</t:Name>
                        <t:EmailAddress>alex.r@company.com</t:EmailAddress>
                      </t:Mailbox>
                    </t:Sender>
                    <t:Body BodyType="Text">This looks stunning, Sarah! The micro-animations on hover and the dark mode gradients feel extremely premium. Performance-wise, we just need to ensure the backdrop-filter styles don't bottleneck mobile rendering. I'll create a sandbox mockup to test.</t:Body>
                  </t:Message>
                </t:Items>
              </t:ConversationNode>
            </t:ConversationNodes>
          </m:Conversation>
        </m:GetConversationItemsResponseMessage>
      </m:ResponseMessages>
    </m:GetConversationItemsResponse>
  </s:Body>
</s:Envelope>`;

// Office Init
Office.onReady((info) => {
    isOfficeInitialized = true;
    const statusBadge = document.getElementById("app-status");
    
    if (info.host === Office.HostType.Outlook) {
        statusBadge.textContent = "Outlook Host";
        statusBadge.className = "status-badge connected";
        initializeAddIn();
    } else {
        // Outside Outlook (e.g. general web browser preview)
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
        });
    });

    // Action buttons
    document.getElementById("btn-copy-json").addEventListener("click", copyJsonToClipboard);
    document.getElementById("btn-download-json").addEventListener("click", downloadJsonFile);
    document.getElementById("btn-copy-xml").addEventListener("click", copyXmlToClipboard);
    document.getElementById("btn-collapse-all").addEventListener("click", () => renderJsonView(rawConversationData, false));
    document.getElementById("btn-expand-all").addEventListener("click", () => renderJsonView(rawConversationData, true));
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

        fetchConversationHistory(conversationId);
    } catch (e) {
        showError(`Initialization error: ${e.message}`);
    }
}

// Fetch via EWS SOAP
function fetchConversationHistory(conversationId) {
    const progressBar = document.getElementById("progress-bar");
    progressBar.style.width = "40%";

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
      <m:Conversations>
        <t:ConversationRequest>
          <t:ConversationId Id="${conversationId}" />
        </t:ConversationRequest>
      </m:Conversations>
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;

    progressBar.style.width = "70%";

    Office.context.mailbox.makeEwsRequestAsync(soapRequest, (asyncResult) => {
        progressBar.style.width = "100%";
        
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            rawXmlResponse = asyncResult.value;
            parseEwsSoapResponse(rawXmlResponse);
        } else {
            showError(`Exchange Server SOAP Call failed: ${asyncResult.error.message}`);
        }
    });
}

// Parse EWS XML response
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
            showError(faultString);
            return;
        }

        // Get conversation nodes
        const conversationNodes = xmlDoc.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "ConversationNode") ||
                                  xmlDoc.getElementsByTagName("t:ConversationNode") ||
                                  xmlDoc.getElementsByTagName("ConversationNode");

        if (conversationNodes.length === 0) {
            showError("No conversation history node structure found in response.");
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
                    participantsSet.add(`${senderName} (${senderEmail})`);
                }

                // Get body content
                let bodySnippet = getXmlNodeText(item, "Body");
                if (bodySnippet && bodySnippet.length > 500) {
                    bodySnippet = bodySnippet.substring(0, 500) + "...";
                }

                if (!mainSubject && subject) {
                    mainSubject = subject;
                }

                messages.push({
                    messageId: itemId || `MSG_${messages.length + 1}`,
                    subject: subject || "No Subject",
                    sender: {
                        name: senderName,
                        email: senderEmail
                    },
                    dateTimeReceived: dateTimeReceived || new Date().toISOString(),
                    bodySnippet: bodySnippet || ""
                });
            }
        }

        // Sort chronological
        messages.sort((a, b) => new Date(a.dateTimeReceived) - new Date(b.dateTimeReceived));

        rawConversationData = {
            conversationId: Office.context.mailbox.item.conversationId,
            subject: mainSubject || Office.context.mailbox.item.subject || "Subject Unavailable",
            totalMessages: messages.length,
            participants: Array.from(participantsSet),
            timeframe: getFormattedTimeframe(messages),
            messages: messages
        };

        renderDashboard(rawConversationData, xmlString);
    } catch (e) {
        showError(`Parsing XML response failed: ${e.message}`);
    }
}

// Utility to get XML text content handling namespaces
function getXmlNodeText(parentNode, localName) {
    const nodes = parentNode.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", localName);
    if (nodes.length > 0) return nodes[0].textContent;
    
    const prefNodes = parentNode.getElementsByTagName(`t:${localName}`);
    if (prefNodes.length > 0) return prefNodes[0].textContent;
    
    const plainNodes = parentNode.getElementsByTagName(localName);
    if (plainNodes.length > 0) return plainNodes[0].textContent;
    
    return "";
}

// Utility to get XML attributes handling namespaces
function getXmlNodeValue(parentNode, localName, attributeName) {
    const nodes = parentNode.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", localName);
    if (nodes.length > 0) return nodes[0].getAttribute(attributeName);
    
    const prefNodes = parentNode.getElementsByTagName(`t:${localName}`);
    if (prefNodes.length > 0) return prefNodes[0].getAttribute(attributeName);
    
    const plainNodes = parentNode.getElementsByTagName(localName);
    if (plainNodes.length > 0) return plainNodes[0].getAttribute(attributeName);
    
    return "";
}

// Helper to format conversation timeframe
function getFormattedTimeframe(messages) {
    if (messages.length === 0) return "";
    if (messages.length === 1) {
        return formatDate(messages[0].dateTimeReceived);
    }
    const firstDate = formatDate(messages[0].dateTimeReceived);
    const lastDate = formatDate(messages[messages.length - 1].dateTimeReceived);
    return `${firstDate} - ${lastDate}`;
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
function renderDashboard(data, xmlResponse = "") {
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

    // Populate Raw XML debug view
    document.getElementById("xml-output").textContent = xmlResponse || "XML schema unavailable in Preview Mode.";
}

// Renders the visual Timeline tab
function renderTimelineView(messages) {
    const timelineList = document.getElementById("timeline-list");
    timelineList.innerHTML = "";

    messages.forEach((msg, index) => {
        const item = document.createElement("li");
        item.className = "timeline-item";
        item.style.animationDelay = `${index * 0.1}s`;

        const d = new Date(msg.dateTimeReceived);
        const dateStr = d.toLocaleDateString("en-US", { month: "short", day: "numeric" }) + " " + 
                      d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });

        item.innerHTML = `
            <div class="timeline-marker"></div>
            <div class="timeline-content">
                <div class="timeline-header">
                    <span class="sender-name">${escapeHtml(msg.sender.name || msg.sender.email)}</span>
                    <span class="date-text">${dateStr}</span>
                </div>
                <div class="snippet-text">${escapeHtml(msg.bodySnippet || "No content summary available.")}</div>
            </div>
        `;
        timelineList.appendChild(item);
    });
}

// Interactive Collapsible JSON Renderer
function renderJsonView(obj, defaultExpanded = true) {
    const jsonOutput = document.getElementById("json-output");
    jsonOutput.innerHTML = "";
    
    const html = formatJsonHtml(obj, defaultExpanded);
    jsonOutput.appendChild(html);
}

function formatJsonHtml(val, expanded) {
    const container = document.createElement("div");
    
    if (val === null) {
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

        const block = document.createElement("span");
        const openBracket = document.createElement("span");
        openBracket.className = "collapsible-trigger";
        openBracket.textContent = expanded ? "[" : "[ ... ]";
        
        const closeBracket = document.createElement("span");
        closeBracket.textContent = "]";
        
        const contents = document.createElement("div");
        contents.className = "json-node";
        if (!expanded) contents.style.display = "none";

        openBracket.onclick = () => {
            if (contents.style.display === "none") {
                contents.style.display = "block";
                openBracket.textContent = "[";
            } else {
                contents.style.display = "none";
                openBracket.textContent = "[ ... ]";
            }
        };

        val.forEach((item, index) => {
            const itemDiv = document.createElement("div");
            itemDiv.appendChild(formatJsonHtml(item, expanded));
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

        const block = document.createElement("span");
        const openBrace = document.createElement("span");
        openBrace.className = "collapsible-trigger";
        openBrace.textContent = expanded ? "{" : "{ ... }";
        
        const closeBrace = document.createElement("span");
        closeBrace.textContent = "}";
        
        const contents = document.createElement("div");
        contents.className = "json-node";
        if (!expanded) contents.style.display = "none";

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
            rowDiv.appendChild(formatJsonHtml(val[key], expanded));
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

function copyXmlToClipboard() {
    if (!rawXmlResponse) return;
    navigator.clipboard.writeText(rawXmlResponse).then(() => {
        const btn = document.getElementById("btn-copy-xml");
        btn.textContent = "Copied ✓";
        setTimeout(() => {
            btn.textContent = "Copy XML";
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
    a.download = `outlook-conversation-${rawConversationData.conversationId || "thread"}.json`;
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
    rawConversationData = mockConversation;
    rawXmlResponse = mockXmlResponse;
    renderDashboard(rawConversationData, rawXmlResponse);
}
