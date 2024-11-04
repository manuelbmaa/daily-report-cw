// index.js
const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs'); // Importa correctamente ExcelJS
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PRIORITY_USERS = [
  "ABRAHAM SANCHEZ",
  "CARLOS EDUARDO NIETO MORALES",
  "DAYANA MENDEZ ANDRADE",
  "DIEGO BARRIOS",
  "DIXINY YRAIDA ACERO ROSALES",
  "FRANCISCO JAVIER TOUCEIRO RODRIGUEZ",
  "GABRIEL GRIMALDI",
  "GIORGIO MARINETTI",
  "HERBERTH ALEXIS PORTILLO MONTES",
  "ARIEL ANTONIO CHACON RIVERA",
  "MARIANA TINEO",
  "MOISES CARVALHO PETIT",
  "NELSON BOLIVAR",
  "OLGA MAIRIETTE CERÓN TORRES",
  "PEDRO ALVAREZ",
  "PIERINA GARCIA",
  "RAFAEL ANDRES PANFIL VILLANUEVA",
  "SANTY JESÚS BEVILACQUA",
  "SERGIO ALFONSO PAZOS NOGUERA",
  "SONIA ESPINOZA",
  "TOMAS HERNANDEZ",
  "VALENTINA CEDENO",
  "VANESSA PASQUALE BONILLA",
  "MANUEL RODRIGUES",
  "PIA CARO",
  "JASMIN ARTEAGA",
  "FELIX CANISALEZ",
  "MIGUEL ANGEL ARVELO",
  "NADIM HANNA",
  "JOSSALVY MILLAN",
];

// Función para calcular la diferencia de tiempo en formato HH:MM:SS
function calculateHoursDifference(start, end) {
  const msDifference = end - start;
  const hours = Math.floor(msDifference / (1000 * 60 * 60));
  const minutes = Math.floor((msDifference % (1000 * 60 * 60)) / (1000 * 60));
  const seconds = Math.floor((msDifference % (1000 * 60)) / 1000);
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

// Función para obtener el nombre del día en español
function getDayName(date) {
  const days = ["DOMINGO", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"];
  return `${days[date.getDay()]} ${String(date.getDate()).padStart(2, '0')}`;
}

// Función para formatear el tiempo en HH:MM:SS
function formatTime(date) {
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${hours}:${minutes}:${seconds}`;
}

app.use(express.static(path.join(__dirname, 'public')));

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    let fileContent = fs.readFileSync(filePath, 'utf-8');

    // Eliminar la primera línea si es un encabezado extra
    const lines = fileContent.split('\n');
    if (lines[0].includes("Eventos de Hoy")) {
      lines.shift();
    }
    fileContent = lines.join('\n');

    // Procesar el CSV ignorando el BOM y comenzar desde la segunda línea
    const records = parse(fileContent, {
      columns: true,
      bom: true,
      skip_empty_lines: true
    });

    // Crear una estructura para agrupar los marcajes por usuario
    const userEntries = {};
    records.forEach(record => {
      const fullName = `${record.Nombre} ${record.Apellido}`;
      if (PRIORITY_USERS.includes(fullName)) {
        if (!userEntries[fullName]) {
          userEntries[fullName] = [];
        }
        userEntries[fullName].push(new Date(record.Tiempo));
      }
    });

    // Generar el archivo Excel con formato
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Report");

    // Obtener el nombre del día desde el primer registro de fecha
    const firstDate = new Date(records[0].Tiempo);
    const dayName = getDayName(firstDate);

    // Configurar las columnas con el día de la semana
    worksheet.columns = [
      { header: "NOMBRE", key: "name", width: 30 },
      { header: "REPORTE", key: "report", width: 20 },
      { header: dayName, key: "time", width: 20 },
    ];

    // Estilo del encabezado (aplicar color azul y fuente Calibri 9 en toda la fila)
    worksheet.getRow(1).eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '002060' },
      };
      cell.font = { name: 'Calibri', size: 9, color: { argb: 'FFFFFF' }, bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // Añadir los datos de cada usuario en el orden de PRIORITY_USERS
    let rowIndex = 2;
    PRIORITY_USERS.forEach(name => {
      const times = userEntries[name];
      if (times) {  // Solo si hay registros para el usuario
        // Ordenar las horas para obtener el primer y último marcaje
        times.sort((a, b) => a - b);
        const firstEntry = times[0];
        const lastEntry = times[times.length - 1];

        const firstEntryTime = formatTime(firstEntry);
        const lastEntryTime = formatTime(lastEntry);

        // Calcular las horas reportadas
        const hoursReported = calculateHoursDifference(firstEntry, lastEntry);

        // Combinar y centrar el nombre del usuario en tres filas
        worksheet.mergeCells(`A${rowIndex}:A${rowIndex + 2}`);
        const nameCell = worksheet.getCell(`A${rowIndex}`);
        nameCell.value = name;
        nameCell.alignment = { vertical: 'middle', horizontal: 'center' };
        nameCell.font = { name: 'Calibri', size: 9, bold: true }; // Aplica negrita aquí

        // Aplicar bordes y centrado a cada celda en el bloque de tres filas
        for (let i = 0; i < 3; i++) {
          worksheet.getCell(`A${rowIndex + i}`).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
          worksheet.getCell(`B${rowIndex + i}`).font = { name: 'Calibri', size: 9 };
          worksheet.getCell(`B${rowIndex + i}`).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
          worksheet.getCell(`C${rowIndex + i}`).font = { name: 'Calibri', size: 9 };
          worksheet.getCell(`C${rowIndex + i}`).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
          worksheet.getCell(`C${rowIndex + i}`).alignment = {
            vertical: 'middle',
            horizontal: 'center'
          };
        }

        // Primer Marcaje
        worksheet.getCell(`B${rowIndex}`).value = "Primer Marcaje";
        worksheet.getCell(`C${rowIndex}`).value = firstEntryTime;

        // Último Marcaje
        worksheet.getCell(`B${rowIndex + 1}`).value = "Último Marcaje";
        worksheet.getCell(`C${rowIndex + 1}`).value = lastEntryTime;

        // Horas Reportadas (con color de fondo azul en columna B y verde en columna C)
        worksheet.getCell(`B${rowIndex + 2}`).value = "Horas Reportadas";
        worksheet.getCell(`B${rowIndex + 2}`).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '002060' },
        };
        worksheet.getCell(`B${rowIndex + 2}`).font = { name: 'Calibri', size: 9, color: { argb: 'FFFFFF' }, bold: true };
        worksheet.getCell(`B${rowIndex + 2}`).alignment = { vertical: 'middle' };

        worksheet.getCell(`C${rowIndex + 2}`).value = hoursReported;
        worksheet.getCell(`C${rowIndex + 2}`).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '8ED973' }
        };
        worksheet.getCell(`C${rowIndex + 2}`).alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell(`C${rowIndex + 2}`).font = { name: 'Calibri', size: 9, bold: true };
        
        // Incrementar el índice para el próximo usuario sin dejar filas en blanco
        rowIndex += 3;
      }
    });

    // Guardar y enviar el archivo Excel como respuesta
    const fileName = 'Report.xlsx';
    const filePathToSave = path.join(__dirname, fileName);
    await workbook.xlsx.writeFile(filePathToSave);

    // Descargar el archivo y luego eliminarlo
    res.download(filePathToSave, fileName, (err) => {
      if (err) throw err;
      fs.unlinkSync(filePath); // Elimina el archivo CSV cargado
      fs.unlinkSync(filePathToSave); // Elimina el archivo Excel generado
    });
  } catch (error) {
    console.error("Error processing file upload:", error);
    res.status(500).send("Error processing file upload");
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Servidor corriendo en el puerto ${PORT}`));
