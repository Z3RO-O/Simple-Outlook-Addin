Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("btn").onclick = run;
    document.getElementById("file-upload").onchange = previewImage;
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

    if (item.body) {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        loadingIndicator.style.display = "none";
        buttonViewWhileLoading.style.display = "block";

        if (result.status === Office.AsyncResultStatus.Succeeded) {
          bodyContainer.textContent = result.value || "No body content";
        } else {
          bodyContainer.textContent = "Unable to retrieve body content";
        }
      });
    } else {
      bodyContainer.textContent = "No body content";
      loadingIndicator.style.display = "none";
      buttonViewWhileLoading.style.display = "block";
    }
  } catch (error) {
    console.error("Error extracting email data:", error);
    loadingIndicator.style.display = "none";
  }
}

function previewImage(event) {
  const file = event.target.files[0];
  const preview = document.getElementById("image-preview");

  if (file && file.type.startsWith("image/")) {
    const reader = new FileReader();

    reader.onload = (e) => {
      preview.src = e.target.result;
      preview.style.display = "block";
    };

    reader.readAsDataURL(file);
  } else {
    alert("Please upload a valid image file.");
    preview.style.display = "none";
  }
}