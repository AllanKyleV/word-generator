const nameInput = document.getElementById("name");
const messageInput = document.getElementById("message");
const generateBtn = document.getElementById("generateBtn");

generateBtn.addEventListener("click", async () => {
  // 1️⃣ Load the template
  const response = await fetch("template.docx");
  const arrayBuffer = await response.arrayBuffer();

  // 2️⃣ Create zip from template
  const zip = new PizZip(arrayBuffer);

  // 3️⃣ Create Docxtemplater instance
  const doc = new window.Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  // 4️⃣ Set the data for placeholders
  doc.setData({
    name: nameInput.value,
    message: messageInput.value
  });

  try {
    // 5️⃣ Render the document
    doc.render();
  } catch (error) {
    console.error(error);
    return;
  }

  // 6️⃣ Generate the blob
  const out = doc.getZip().generate({ type: "blob" });

  // 7️⃣ Trigger download
  const link = document.createElement("a");
  link.href = URL.createObjectURL(out);
  link.download = "generated.docx";
  link.click();
});
