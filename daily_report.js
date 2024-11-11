// daily_report.js

const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Sirve archivos estáticos desde la carpeta "public"
app.use(express.static(path.join(__dirname, 'public')));

const PRIORITY_USERS = [
  "ABRAHAM SANCHEZ",
  "CARLOS EDUARDO NIETO MORALES",
  "DAYANA MENDEZ ANDRADE",
  "DIEGO BARRIOS",
  "DIXINY ACERO",
  "FRANCISCO JAVIER TOUCEIRO RODRIGUEZ",
  "GABRIEL GRIMALDI",
  "GIORGIO MARINETTI",
  "HERBERTH ALEXIS PORTILLO MONTES",
  "JONATHAN CHACÓN",
  "MARIANA TINEO",
  "MOISES CARVALHO",
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
  "MIGUEL ARVELO",
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
  const days = ["DOMINGO", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"];
  return `${days[date.getDay()]} ${String(date.getDate()).padStart(2, '0')}`;
}

// Función para formatear el tiempo en HH:MM:SS
function formatTime(date) {
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${hours}:${minutes}:${seconds}`;
}

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    let fileContent = fs.readFileSync(filePath, 'utf-8');

    const lines = fileContent.split('\n');
    if (lines[0].includes("Eventos de Hoy")) {
      lines.shift();
    }
    fileContent = lines.join('\n');

    const records = parse(fileContent, {
      columns: true,
      bom: true,
      skip_empty_lines: true
    });

    const userEntries = {};
    records.forEach(record => {
      const fullName = `${record.Nombre} ${record.Apellido}`;
      if (PRIORITY_USERS.includes(fullName)) {
        if (!userEntries[fullName]) {
          userEntries[fullName] = [];
        }
        const entryTime = new Date(record.Tiempo);
        if (!isNaN(entryTime)) {
          userEntries[fullName].push(entryTime);
        }
      }
    });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Report");

    const firstDate = new Date(records[0].Tiempo); // Obtén la fecha del primer registro
    const dayName = getDayName(firstDate);

    worksheet.columns = [
      { header: "NOMBRE", key: "name", width: 30 },
      { header: "REPORTE", key: "report", width: 20 },
      { header: dayName, key: "time", width: 20 },
    ];

    worksheet.getRow(1).eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '002060' },
      };
      cell.font = { name: 'Calibri Light', size: 9, color: { argb: 'FFFFFF' }, bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    let rowIndex = 2;
    PRIORITY_USERS.forEach(name => {
      const times = userEntries[name];
      if (times && times.length > 0) {
        times.sort((a, b) => a - b);
        const firstEntry = times[0];
        const lastEntry = times[times.length - 1];

        const firstEntryTime = formatTime(firstEntry);
        const lastEntryTime = formatTime(lastEntry);
        const hoursReported = calculateHoursDifference(firstEntry, lastEntry);

      // Escribir en las celdas en el orden de PRIORITY_USERS
      worksheet.mergeCells(`A${rowIndex}:A${rowIndex + 2}`);
      const nameCell = worksheet.getCell(`A${rowIndex}`);
      nameCell.value = name;
      nameCell.alignment = { vertical: 'middle', horizontal: 'center' };
      nameCell.font = { name: 'Calibri Light', size: 9, bold: true };

      for (let i = 0; i < 3; i++) {
        worksheet.getCell(`A${rowIndex + i}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      worksheet.getCell(`B${rowIndex + i}`).font = { name: 'Calibri Light', size: 9 };
      worksheet.getCell(`B${rowIndex + i}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      worksheet.getCell(`C${rowIndex + i}`).font = { name: 'Calibri Light', size: 9 };
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

    worksheet.getCell(`B${rowIndex}`).value = "Primer Marcaje";
    worksheet.getCell(`C${rowIndex}`).value = firstEntryTime;

    worksheet.getCell(`B${rowIndex + 1}`).value = "Último Marcaje";
    worksheet.getCell(`C${rowIndex + 1}`).value = lastEntryTime;

    worksheet.getCell(`B${rowIndex + 2}`).value = "Horas Reportadas";
    worksheet.getCell(`B${rowIndex + 2}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '002060' },
    };
    worksheet.getCell(`B${rowIndex + 2}`).font = { name: 'Calibri Light', size: 9, color: { argb: 'FFFFFF' }, bold: true };
    worksheet.getCell(`B${rowIndex + 2}`).alignment = { vertical: 'middle' };

    worksheet.getCell(`C${rowIndex + 2}`).value = hoursReported;
    worksheet.getCell(`C${rowIndex + 2}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '8ED973' }
    };
    worksheet.getCell(`C${rowIndex + 2}`).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell(`C${rowIndex + 2}`).font = { name: 'Calibri Light', size: 9, bold: true };
    
    rowIndex += 3;
  }
});

    // Formatear la fecha del primer registro como "dd-MM-yyyy"
    const formattedDate = `${String(firstDate.getDate()).padStart(2, '0')}-${String(firstDate.getMonth() + 1).padStart(2, '0')}-${firstDate.getFullYear()}`;

    // Nombre del archivo con la fecha del primer registro
    const fileName = `MARCAJE DEL DIA CW ${formattedDate}.xlsx`;
    const filePathToSave = path.join(__dirname, fileName);
    
    await workbook.xlsx.writeFile(filePathToSave);

    res.download(filePathToSave, fileName, (err) => {
      if (err) throw err;
      fs.unlinkSync(filePath);
      fs.unlinkSync(filePathToSave);
    });
  } catch (error) {
    console.error("Error processing file upload:", error);
    res.status(500).send("Error processing file upload");
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Servidor corriendo en el puerto ${PORT}`));
