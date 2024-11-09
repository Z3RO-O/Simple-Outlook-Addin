Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("btn").onclick = run;
    // document.getElementById("file-upload").onchange = previewImage;
  }
});

async function run() {
  const item = Office.context.mailbox.item;

  const loadingIndicator = document.getElementById("loading");
  loadingIndicator.style.display = "inline-block";
  const buttonViewWhileLoading = document.getElementById("btn");
  buttonViewWhileLoading.style.display = "none";

  try {
    const subjectContainer = document.getElementById("subject");
    subjectContainer.textContent = item.subject || "No subject";

    const bodyContainer = document.getElementById("body");

    await Promise.all([
      retrieveEmailBody(item, bodyContainer),
      retrieveAttachments(item)
    ]);

    loadingIndicator.style.display = "none";
    buttonViewWhileLoading.style.display = "block";
  } catch (error) {
    console.error("Error extracting email data:", error);
    loadingIndicator.style.display = "none";
    buttonViewWhileLoading.style.display = "block";
  }
}

function retrieveEmailBody(item, bodyContainer) {
  return new Promise((resolve, reject) => {
    if (item.body) {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          bodyContainer.textContent = result.value || "No body content";
          resolve();
        } else {
          bodyContainer.textContent = "Unable to retrieve body content";
          reject("Failed to retrieve body content");
        }
      });
    } else {
      bodyContainer.textContent = "No body content";
      resolve();
    }
  });
}

function retrieveAttachments(item) {
  return new Promise((resolve, reject) => {
    const attachments = item.attachments;

    if (attachments && attachments.length > 0) {
      const imageAttachments = attachments.filter(att => att.contentType && att.contentType.startsWith("image/"));

      if (imageAttachments.length > 0) {
        // Process each image attachment
        const attachmentPromises = imageAttachments.map(att => 
          item.getAttachmentContentAsync(att.id, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
              displayAttachment(result.value.content, result.value.format);
            }
          })
        );

        // Wait until all image attachments are processed
        Promise.all(attachmentPromises).then(resolve).catch(reject);
      } else {
        resolve(); // No image attachments to process
      }
    } else {
      resolve(); // No attachments in the email
    }
  });
}

function displayAttachment(content, format) {
  const preview = document.getElementById("image-preview");
  preview.style.display = "block";

  // Format can be "base64" or "url"
  if (format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    preview.src = `data:image/png;base64,${content}`;
  } else if (format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    preview.src = content;
  }
}