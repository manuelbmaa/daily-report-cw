// Manejo del formulario para el Reporte del Dia
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
            document.getElementById("message").textContent = "Archivo procesado exitosamente. El reporte del día se descargará automáticamente.";
            
            // Obtener el nombre del archivo desde el encabezado Content-Disposition
            const contentDisposition = response.headers.get("Content-Disposition");
            const fileName = contentDisposition
                ? contentDisposition.match(/filename="(.+)"/)[1] // Extrae el nombre del archivo del encabezado
                : "Reporte.xlsx"; // Valor por defecto si no se encuentra

            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = downloadUrl;
            a.download = fileName; // Usa el nombre de archivo obtenido dinámicamente
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

// Manejo del formulario para el Reporte de la Mañana
document.getElementById("morningForm").addEventListener("submit", async (event) => {
    event.preventDefault();

    const formData = new FormData();
    const fileInput = document.getElementById("morningFileInput");
    formData.append("file", fileInput.files[0]);

    try {
        const response = await fetch("/upload_morning", {
            method: "POST",
            body: formData,
        });

        if (response.ok) {
            document.getElementById("morningMessage").textContent = "Archivo procesado exitosamente. El reporte de la mañana se descargará automáticamente.";

            const contentDisposition = response.headers.get("Content-Disposition");
            const fileName = contentDisposition
                ? contentDisposition.match(/filename="(.+)"/)[1]
                : "Morning_Report.xlsx";

            const blob = await response.blob();
            const downloadUrl = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = downloadUrl;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            a.remove();
        } else {
            document.getElementById("morningMessage").textContent = "Error al procesar el archivo.";
        }
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("morningMessage").textContent = "Error al enviar el archivo.";
    }
});