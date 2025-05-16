import createLoader from './loader.js';

let originalEmailBody = "";
let pollingInterval = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    
    // Start polling only on the index page
    if (window.location.pathname.endsWith("/index.html")) {
      
      startEmailPolling();
    }

    // Optional: Monitor URL changes (for SPA-style apps)
    const observer = new MutationObserver(() => {
      if (window.location.pathname.endsWith("/index.html")) {
        if (!pollingInterval) startEmailPolling();
      } else {
        stopEmailPolling();
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  }
});

function startEmailPolling() {
  pollingInterval = setInterval(() => {
    try {
      Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Text,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const currentBody = result.value;
            if (currentBody !== originalEmailBody) {
              originalEmailBody = currentBody;
       
              onEmailChanged();
            }
          } else {
            console.warn("Polling: Failed to get email body:", result.error.message);
          }
        }
      );
    } catch (err) {
      console.error("Polling error:", err);
    }
  }, 2000); // Poll every 2 seconds
}

function stopEmailPolling() {
  if (pollingInterval) {
    clearInterval(pollingInterval);
    pollingInterval = null;
  }
}

function onEmailChanged() {
const startLoading = createLoader(); // show the loader
  getEmailAnalysis()
    .then((replyData) => {
      if (replyData) {
        renderEmailAnalysis(replyData);
      }
    })
    .catch((err) => console.error("Analysis failed:", err))
    .finally(() => {
      const stopLoading = createLoader();
      stopLoading(); // hide the loader
    });
}

async function getEmailAnalysis() {
  const data = { emailText: originalEmailBody };

  try {
    const response = await fetch(
      "https://prod-37.westus.logic.azure.com:443/workflows/3a0de62053ca46169b669d6abcf4fc8d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=usG50jhh1f4Z1578OhwqoR5kOIcYyg1GKMWKmD2zROc",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      }
    );

    const result = await response.json();
    console.log("Response:", result);

    if (result.success) {
      return result.reply; // let the caller handle rendering
    } else {
      console.error("Error in response:", result.error);
      return null;
    }
  } catch (error) {
    console.error("Error fetching data:", error);
    return null;
  }
}

 function renderEmailAnalysis(data) {
    const container = document.getElementById("card-deck");
    container.innerHTML = ""; // Clear existing content
    
    // === Sentiment Card ===
    const sentimentColors = {
      Positive: "bg-green",
      Negative: "bg-red",
      Neutral: "bg-yellow"
    };
    
    const sentimentIcon = {
      
      Positive: `<svg  xmlns="http://www.w3.org/2000/svg"  width="24"  height="24"  viewBox="0 0 24 24"  fill="none"  stroke="currentColor"  stroke-width="2"  stroke-linecap="round"  stroke-linejoin="round"  class="icon icon-tabler icons-tabler-outline icon-tabler-mood-smile"><path stroke="none" d="M0 0h24v24H0z" fill="none"/><path d="M12 12m-9 0a9 9 0 1 0 18 0a9 9 0 1 0 -18 0" /><path d="M9 10l.01 0" /><path d="M15 10l.01 0" /><path d="M9.5 15a3.5 3.5 0 0 0 5 0" /></svg>`,
      Negative: `<svg  xmlns="http://www.w3.org/2000/svg"  width="24"  height="24"  viewBox="0 0 24 24"  fill="none"  stroke="currentColor"  stroke-width="2"  stroke-linecap="round"  stroke-linejoin="round"  class="icon icon-tabler icons-tabler-outline icon-tabler-mood-sad"><path stroke="none" d="M0 0h24v24H0z" fill="none"/><path d="M12 12m-9 0a9 9 0 1 0 18 0a9 9 0 1 0 -18 0" /><path d="M9 10l.01 0" /><path d="M15 10l.01 0" /><path d="M9.5 15.25a3.5 3.5 0 0 1 5 0" /></svg>`,
      Neutral: `<svg  xmlns="http://www.w3.org/2000/svg"  width="24"  height="24"  viewBox="0 0 24 24"  fill="none"  stroke="currentColor"  stroke-width="2"  stroke-linecap="round"  stroke-linejoin="round"  class="icon icon-tabler icons-tabler-outline icon-tabler-mood-empty"><path stroke="none" d="M0 0h24v24H0z" fill="none"/><path d="M12 12m-9 0a9 9 0 1 0 18 0a9 9 0 1 0 -18 0" /><path d="M9 10l.01 0" /><path d="M15 10l.01 0" /><path d="M9 15l6 0" /></svg>`
    }
    
    const sentimentCard = document.createElement("div");
    sentimentCard.className = "card";
    sentimentCard.innerHTML = `
    <div class="card-body">
      <div class="row align-items-center">
        <div class="col-auto">
          <span class="${sentimentColors[data.sentiment] || "bg-gray"} text-white avatar">
            ${sentimentIcon[data.sentiment] || "" }
            </span>
            </div>
            <div class="col">
              <div class="font-weight-medium">${data.sentiment}</div>
              <div class="text-secondary">Email Sentiment</div>
              </div>
              </div>
              </div>
              `;
              container.appendChild(sentimentCard);
              
                  // === Summary Card ===
                  const summaryCard = document.createElement("div");
                  summaryCard.className = "card";
                  summaryCard.innerHTML = `
                    <div class="card-body">
                      <div class="col">
                        <div class="mb-1 col">Summary</div>
                        <div class="text-secondary w-100">${data.summary}</div>
                      </div>
                    </div>
                  `;
                  container.appendChild(summaryCard);
              
    // === Quick Reply Card ===
    const quickReplyCard = document.createElement("div");
    quickReplyCard.className = "card";

    const replies = Array.isArray(data.responses)
      ? data.responses.map(
          r => `<span class="list-group-item text-secondary rounded p-2" aria-current="true">${r.item || ''}</span>`
        ).join("")
      : '<span class="text-secondary">No responses available</span>';

    quickReplyCard.innerHTML = `
      <div class="card-body">
        <div class="mb-1 col">Quick Reply</div>
        <div class="list-group list-group-flush">
          ${replies}
        </div>
      </div>
      <div class="d-flex"></div>
    `;
    container.appendChild(quickReplyCard);
  }


    // window.onload = () =>
    //   document.getElementById("loadingSpinner").classList.remove("show");

    // // Show/hide custom fields
    // const showOtherInput = (selectId, inputId) => {
    //   document
    //     .getElementById(selectId)
    //     .addEventListener("change", function () {
    //       const input = document.getElementById(inputId);
    //       if (this.value === "Other") {
    //         input.classList.remove("d-none");
    //         input.required = true;
    //       } else {
    //         input.classList.add("d-none");
    //         input.required = false;
    //         input.value = "";
    //       }
    //     });
    // };

    // showOtherInput("emailAudience", "audienceOther");
    // showOtherInput("emailTone", "toneOther");
    // showOtherInput("emailType", "typeOther");

    // let templateContent = "";
    // const templateDiv = document.getElementById("templateContent");
    // templateDiv.addEventListener("paste", (event) => {
    //   event.preventDefault();
    //   templateContent = event.clipboardData.getData("text/plain");
    //   templateDiv.innerHTML = event.clipboardData.getData("text/html");
    // });

    // document
    //   .getElementById("replyForm")
    //   .addEventListener("submit", async function (e) {
    //     e.preventDefault();

    //     const formContainer = document.getElementById("formContainer");
    //     const loadingSpinner = document.getElementById("loadingSpinner");
    //     const previewEmail = document.getElementById("emailPreview");

    //     formContainer.classList.add("blurred");
    //     loadingSpinner.classList.add("show");

    //     const emailAudience =
    //       document.getElementById("emailAudience").value === "Other"
    //         ? document.getElementById("audienceOther").value
    //         : document.getElementById("emailAudience").value;

    //     const emailTone =
    //       document.getElementById("emailTone").value === "Other"
    //         ? document.getElementById("toneOther").value
    //         : document.getElementById("emailTone").value;

    //     const emailType =
    //       document.getElementById("emailType").value === "Other"
    //         ? document.getElementById("typeOther").value
    //         : document.getElementById("emailType").value;

    //     const emailDescription =
    //       document.getElementById("emailDescription").value;

    //     const replyPrompt = {
    //       audience: emailAudience,
    //       tone: emailTone,
    //       type: emailType,
    //       description: emailDescription,
    //       template: templateContent,
    //       originalEmail: originalEmailBody,
    //     };

    //     try {
    //       const response = await fetch(
    //         "https://prod-145.westus.logic.azure.com:443/workflows/c65d371e274a4a29a0daa3ee25ce63bd/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-skVHfpWefHb6DNSYMWTmkI3mEUObkxnK2wAvBwU-wg",
    //         {
    //           method: "POST",
    //           headers: { "Content-Type": "application/json" },
    //           body: JSON.stringify(replyPrompt),
    //         }
    //       );

    //       const result = await response.json();

    //       if (result.success && result.reply) {
    //         loadingSpinner.classList.remove("show");
    //         formContainer.classList.remove("active", "blurred");
    //         previewEmail.classList.add("active");

    //         if (Office.context.roamingSettings.get("includeSignature")) {
    //           const signature =
    //             Office.context.roamingSettings.get("userSignature") ||
    //             result.reply.signature;
    //           const newBody = result.reply.body + "<br>" + signature;
    //           result.reply.body = newBody;
    //         }

    //         previewEmail.innerHTML = `
    //         <div class="card-body fade-in">
    //             <h2 class="mb-4">Generated Reply</h2>
    //             <label class="mb-1">Subject:</label>
    //             <div class="card mb-3">
    //               <div class="card-body overflow-scroll">
    //                 ${result.reply.subject}
    //               </div>
    //             </div>
    //             <label class="mb-1">Body:</label>
    //             <div class="card mb-3">
    //               <div class="card-body overflow-scroll">
    //                 ${result.reply.body}
    //               </div>
    //             </div>
    //           <div class="card-footer bg-transparent mt-auto">
    //           <div class="btn-list justify-content-end">
    //             <a href="#" class="btn btn-1" id="cancelReplyButton">Cancel</a>
    //             <a href="#" class="btn btn-success btn-2" id="insertReplyButton">Insert into Email</a>
    //           </div>
    //         </div>
    //         </div>
    //       `;

    //         document
    //           .getElementById("insertReplyButton")
    //           .addEventListener("click", function () {
    //             const item = Office.context.mailbox.item;

    //             if (
    //               Office.context.mailbox.item.itemType ===
    //               Office.MailboxEnums.ItemType.Message &&
    //               typeof item.body.setSelectedDataAsync === "function"
    //             ) {
    //               // Compose mode: insert text directly
    //               item.body.setSelectedDataAsync(
    //                 insertReplyHtml,
    //                 { coercionType: Office.CoercionType.Text },
    //                 function (asyncResult) {
    //                   if (
    //                     asyncResult.status ===
    //                     Office.AsyncResultStatus.Succeeded
    //                   ) {
    //                     console.log("Reply inserted!");
    //                   } else {
    //                     console.error(
    //                       "Failed to insert reply:",
    //                       asyncResult.error.message
    //                     );
    //                   }
    //                 }
    //               );
    //             } else {
    //               // Not in compose mode, launch new message
    //               Office.context.mailbox.item.displayReplyAllFormAsync(
    //                 {
    //                   htmlBody: result.reply.body,
    //                 },
    //                 function (asyncResult) {
    //                   if (
    //                     asyncResult.status ===
    //                     Office.AsyncResultStatus.Succeeded
    //                   ) {
    //                     console.log("Reply inserted!");
    //                   } else {
    //                     console.error(
    //                       "Failed to insert reply:",
    //                       asyncResult.error.message
    //                     );
    //                   }
    //                 }
    //               );
    //             }
    //             previewEmail.innerHTML = "";
    //             goHome();
    //           });

    //         document
    //           .getElementById("cancelReplyButton")
    //           .addEventListener("click", function () {
    //             previewEmail.innerHTML = "";
    //             goHome();
    //           });

    //         document
    //           .getElementById("emailPreview")
    //           .scrollIntoView({ behavior: "smooth" });
    //       } else {
    //         console.log("Something went wrong. No reply text received.");
    //       }
    //     } catch (err) {
    //       console.error("Fetch error:", err);
    //       console.log("An error occurred while generating the reply.");
    //     }
    //   });

    // function convertEscapedToHtml(input, options = {}) {
    //   if (!input || typeof input !== "string") return "";

    //   if (options.usePre) {
    //     return `<pre>${escapeHtml(input)}</pre>`;
    //   }

    //   return escapeHtml(input)
    //     .replace(/\n/g, "<br>")
    //     .replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;");
    // }

    // function escapeHtml(str) {
    //   return str
    //     .replace(/&/g, "&amp;")
    //     .replace(/</g, "&lt;")
    //     .replace(/>/g, "&gt;")
    //     .replace(/"/g, "&quot;")
    //     .replace(/'/g, "&#39;");
    // }

    // function showSettings() {
    //   document.querySelector(".view.active").classList.remove("active");
    //   document.getElementById("settingsContainer").classList.add("active");
    //   const signatureHtml =
    //     Office.context.roamingSettings.get("userSignature") || "";
    //   document.getElementById("signatureInput").innerHTML = signatureHtml;
    // }

    // function goHome() {
    //   document.querySelector(".view.active").classList.remove("active");
    //   document.getElementById("formContainer").classList.add("active");
    // }

    // function saveSettings() {
    //   const signatureHtml =
    //     document.getElementById("signatureInput").innerHTML;
    //   const signatureSwitch =
    //     document.getElementById("includeSigChk").checked;

    //   Office.context.roamingSettings.set("userSignature", signatureHtml);
    //   Office.context.roamingSettings.set("includeSignature", signatureSwitch);

    //   Office.context.roamingSettings.saveAsync((result) => {
    //     if (result.status === Office.AsyncResultStatus.Succeeded) {
    //       console.log("Settings saved successfully.");
    //       goHome();
    //     } else {
    //       console.log("Failed to save settings.");
    //     }
    //   });

    //   goHome();
    // }