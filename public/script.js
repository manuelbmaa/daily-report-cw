// script.js
document.getElementById("uploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
  
    const formData = new FormData();
    const fileInput = document.getElementById("fileInput");
    formData.append("file", fileInput.files[0]);
  
    try {
      const response = await fetch("/upload", {
        method: "POST",
        body: formData,
      });
  
      if (response.ok) {
        document.getElementById("message").textContent = "Archivo procesado exitosamente. El reporte se descargará automáticamente.";
        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = downloadUrl;
        a.download = "Report.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
      } else {
        document.getElementById("message").textContent = "Error al procesar el archivo.";
      }
    } catch (error) {
      console.error("Error:", error);
      document.getElementById("message").textContent = "Error al enviar el archivo.";
    }
  });
  