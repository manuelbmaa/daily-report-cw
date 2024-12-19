// daily_report.js

const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Importa los arreglos desde el users_mega.js
const { CW_USERS, INJUVE_USERS, MOTORISTAS_USERS, ADDN_USERS, IT_PROYECTOS_USERS, INVEST_USERS, NAME_MAPPING } = require('./users_mega');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Sirve archivos estáticos desde la carpeta "public"
app.use(express.static(path.join(__dirname, 'public')));

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

//--------------------------------------------------------------------------------------------------------------------------------
// Logica del Reporte del Dia
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    const { userGroup } = req.body;  // Ahora debes recibir correctamente el grupo
    
    // Mapeo de grupos a nombres de hojas y archivos
    const groupNames = {
      CW_USERS: "Marcaje del Día CW",
      INJUVE_USERS: "Marcaje del Día INJUVE",
      MOTORISTAS_USERS: "Marcaje del Día MOTORISTAS",
      ADDN_USERS: "Marcaje del Día ADDN",
      IT_PROYECTOS_USERS: "Marcaje del Día IT PROYECTOS",
      INVEST_USERS: "Marcaje del Día INVEST"
    };

    // Validación del grupo de usuarios
    const validGroups = {
      CW_USERS,
      INJUVE_USERS,
      MOTORISTAS_USERS,
      ADDN_USERS,
      IT_PROYECTOS_USERS,
      INVEST_USERS
    };

    const selectedUsers = validGroups[userGroup]; // Obtiene la lista de usuarios correcta
    if (!selectedUsers) {
      throw new Error(`Grupo de usuarios no válido: ${userGroup}`);
    }
    
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
      if (selectedUsers.includes(fullName)) {
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
    // Usa el nombre del grupo para asignar el nombre de la hoja
    const sheetName = groupNames[userGroup] || "Report"; 
    const worksheet = workbook.addWorksheet(sheetName);

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
    // Configurar la celda para el nombre del usuario
    selectedUsers.forEach(name => {
      const times = userEntries[name];
      if (times && times.length > 0) {
        times.sort((a, b) => a - b);
        const firstEntry = times[0];
        const lastEntry = times[times.length - 1];

        const firstEntryTime = formatTime(firstEntry);
        const lastEntryTime = formatTime(lastEntry);
        const hoursReported = calculateHoursDifference(firstEntry, lastEntry);

        // Sustituir el nombre por el preferido si existe en NAME_MAPPING
        const displayName = NAME_MAPPING[name] || name;

        // Escribir en las celdas en el orden de CW_USERS
        worksheet.mergeCells(`A${rowIndex}:A${rowIndex + 2}`);
        const nameCell = worksheet.getCell(`A${rowIndex}`);
        nameCell.value = displayName; // Escribe el nombre preferido
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
    const fileName = `${groupNames[userGroup]} ${formattedDate}.xlsx`;
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

//--------------------------------------------------------------------------------------------------------------------------------
// Logica de Reporte de la Mañana

function formatTimeMorning(date) {
  if (!date) return ""; // Retorna una cadena vacía si el valor es null o undefined
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${hours}:${minutes}:${seconds}`;
}

app.post('/upload_morning', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    let fileContent = fs.readFileSync(filePath, 'utf-8');

    // Filtrar líneas innecesarias
    const lines = fileContent.split('\n').filter(line => {
      return line.trim() !== '' && !line.includes("Eventos de Hoy"); // Ignorar líneas vacías y encabezados no deseados
    });
    fileContent = lines.join('\n');

    // Configurar csv-parse para ignorar errores de longitud inconsistente
    const records = parse(fileContent, {
      columns: true,
      bom: true,
      skip_empty_lines: true,
      relax_column_count: true  // Ignorar filas con número inconsistente de columnas
    });

    const userEntries = {};
    CW_USERS.forEach(name => {
      userEntries[name] = null; // Inicializa con null para cada usuario
    });

    records.forEach(record => {
      const fullName = `${record.Nombre} ${record.Apellido}`;
      if (CW_USERS.includes(fullName)) {
        const entryTime = new Date(record.Tiempo);
        if (!isNaN(entryTime) && (!userEntries[fullName] || entryTime < userEntries[fullName])) {
          userEntries[fullName] = entryTime; // Guarda el primer marcaje más temprano
        }
      }
    });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Morning Report");

    const firstDate = new Date(records[0].Tiempo);
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
    });

    let rowIndex = 2;
    CW_USERS.forEach(name => {
      const displayName = NAME_MAPPING[name] || name; // Usa el nombre mapeado si existe, o el original
      const nameCell = worksheet.getCell(`A${rowIndex}`);
      const reportCell = worksheet.getCell(`B${rowIndex}`);
      const timeCell = worksheet.getCell(`C${rowIndex}`);

      // Configurar el estilo de la celda del nombre
      nameCell.value = displayName; // Escribe el nombre mapeado
      nameCell.font = { name: 'Calibri Light', size: 9, bold: true };

      // Configurar el estilo de la celda "Primer Marcaje"
      reportCell.value = "Primer Marcaje";
      reportCell.font = { name: 'Calibri Light', size: 9 }; // Sin negrita

      // Configurar el estilo de la celda de tiempo
      timeCell.value = formatTimeMorning(userEntries[name]);
      timeCell.font = { name: 'Calibri Light', size: 9 }; // Sin negrita

      worksheet.getRow(rowIndex).eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
      rowIndex++;
    });

    const formattedDate = `${String(firstDate.getDate()).padStart(2, '0')}-${String(firstDate.getMonth() + 1).padStart(2, '0')}-${firstDate.getFullYear()}`;
    const fileName = `REPORTE DE LA MAÑANA CW ${formattedDate}.xlsx`;
    const filePathToSave = path.join(__dirname, fileName);

    await workbook.xlsx.writeFile(filePathToSave);

    res.download(filePathToSave, fileName, (err) => {
      if (err) throw err;
      fs.unlinkSync(filePath);
      fs.unlinkSync(filePathToSave);
    });
  } catch (error) {
    console.error("Error processing morning report:", error);
    res.status(500).send("Error processing morning report");
  }
});

//--------------------------------------------------------------------------------------------------------------------------------
// Logica de Reporte Semanal

// Definir los grupos de usuarios
const userGroups = {
  CW: CW_USERS,
  INJUVE: INJUVE_USERS,
  MOTORISTAS: MOTORISTAS_USERS,
  ADDN: ADDN_USERS,
  "IT PROYECTOS": IT_PROYECTOS_USERS,
  INVEST: INVEST_USERS,
};

// Función para normalizar nombres
function normalizeName(name) {
  return name ? name.trim().toUpperCase() : null;
}

// Función para calcular la diferencia de tiempo en formato HH:MM:SS
function calculateHoursDifference(start, end) {
  if (!start || !end) return "00:00:00";
  const msDifference = end - start;
  const hours = Math.floor(msDifference / (1000 * 60 * 60));
  const minutes = Math.floor((msDifference % (1000 * 60 * 60)) / (1000 * 60));
  const seconds = Math.floor((msDifference % (1000 * 60)) / 1000);
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

// Función para formatear el tiempo en HH:MM:SS
function formatTime(date) {
  if (!date || isNaN(date)) return "00:00:00";
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${hours}:${minutes}:${seconds}`;
}

// Función para verificar si una fecha es válida y pertenece a días laborables
function isWeekday(date) {
  const day = date.getDay();
  return day >= 1 && day <= 5; // Lunes = 1, ..., Viernes = 5
}

// Función para obtener el nombre del día de la semana
function getDayName(date) {
  const dayNames = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES"];
  const dayIndex = date.getDay() - 1; // Ajustar para que lunes sea el índice 0
  if (dayIndex < 0 || dayIndex >= dayNames.length) return null; // Validar días fuera de rango
  return `${dayNames[dayIndex]} ${String(date.getDate()).padStart(2, '0')}`;
}

// Función para actualizar el primer y último marcaje
function updateEntry(entry, entryTime) {
  if (!entry.first || new Date(entryTime) < new Date(entry.first)) {
    entry.first = entryTime;
  }
  if (!entry.last || new Date(entryTime) > new Date(entry.last)) {
    entry.last = entryTime;
  }
}

// Configuración de bordes para las celdas
const borderStyle = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

// Estilo general para todas las celdas
const generalFontStyle = {
  name: 'Calibri Light',
  size: 9,
};

// Estilo para centrado de texto
const centeredAlignment = {
  vertical: 'middle',
  horizontal: 'center'
};

// Función para aplicar bordes a todas las celdas
function applyBordersToRow(row) {
  row.eachCell(cell => {
    cell.border = borderStyle;
  });
}

// Endpoint para procesar y generar el reporte semanal
app.post('/upload_weekly', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    let fileContent = fs.readFileSync(filePath, 'utf-8');

    // Filtrar encabezados innecesarios
    const lines = fileContent.split('\n').filter(line => line.trim() !== '');
    if (lines[0].includes("Todos los Eventos")) {
      lines.shift();
    }
    fileContent = lines.join('\n');

    // Leer y parsear los registros del CSV
    const records = parse(fileContent, {
      columns: true,
      bom: true,
      skip_empty_lines: true,
      relax_column_count: true,
    });

    // Filtrar registros solo de lunes a viernes y con fechas válidas
    const filteredRecords = records.filter(record => {
      const entryTime = new Date(record.Tiempo);
      return isWeekday(entryTime) && !isNaN(entryTime);
    });

    // Extraer fechas únicas y ordenarlas
    const uniqueDates = [...new Set(filteredRecords.map(record => {
      const date = new Date(record.Tiempo);
      return date.toISOString().split('T')[0];
    }))].sort();

    // Convertir fechas únicas a formato legible
    const daysOfWeek = uniqueDates.map(dateStr => {
      const date = new Date(dateStr);
      return getDayName(date);
    }).filter(Boolean); // Filtrar valores nulos o inválidos

    const workbook = new ExcelJS.Workbook();

    // Generar una hoja para cada grupo de usuarios
    for (const [sheetName, users] of Object.entries(userGroups)) {
      const worksheet = workbook.addWorksheet(sheetName);

      // Configurar columnas
      worksheet.columns = [
        { header: "NOMBRE", key: "name", width: 30 },
        { header: "REPORTE", key: "report", width: 20 },
        ...daysOfWeek.map(day => ({ header: day, key: day, width: 20 })),
      ];

      // Estilo de encabezados
      worksheet.getRow(1).eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '002060' },
        };
        cell.font = { ...generalFontStyle, color: { argb: 'FFFFFF' }, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = borderStyle;
      });

      // Inicializar datos de usuarios
      const userEntries = {};
      users.forEach(user => {
        userEntries[user] = uniqueDates.map(() => ({ first: null, last: null }));
      });

      // Procesar registros
      filteredRecords.forEach(record => {
        const fullName = normalizeName(`${record.Nombre} ${record.Apellido}`);
        const entryTime = new Date(record.Tiempo);
        if (!users.map(normalizeName).includes(fullName) || isNaN(entryTime)) return;

        const entryDate = entryTime.toISOString().split('T')[0];
        const dayIndex = uniqueDates.indexOf(entryDate);
        if (dayIndex !== -1) {
          const entry = userEntries[fullName][dayIndex];
          updateEntry(entry, entryTime);
        }
      });

      // Rellenar datos en la hoja
      let rowIndex = 2;
      users.forEach(user => {
        const displayName = NAME_MAPPING[user] || user; // Usa el nombre mapeado si existe, o el original
        const entries = userEntries[user];
        worksheet.mergeCells(`A${rowIndex}:A${rowIndex + 2}`);
        const nameCell = worksheet.getCell(`A${rowIndex}`);
        nameCell.value = displayName; // Escribe el nombre mapeado
        nameCell.alignment = centeredAlignment;
        nameCell.font = { ...generalFontStyle, bold: true };
        nameCell.border = borderStyle;

        uniqueDates.forEach((date, index) => {
          const entry = entries[index];
          const firstTime = entry.first ? formatTime(entry.first) : "00:00:00";
          const lastTime = entry.last ? formatTime(entry.last) : "00:00:00";
          const hoursReported = entry.first && entry.last
            ? calculateHoursDifference(entry.first, entry.last)
            : "00:00:00";

          // Rellenar celdas con los datos
          // Configurar la celda "Primer Marcaje"
          worksheet.getCell(`B${rowIndex}`).value = "Primer Marcaje";
          const firstMarkCell = worksheet.getCell(`B${rowIndex}`);
          firstMarkCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFF' }, // Fondo azul oscuro
          };
          firstMarkCell.font = {
              color: { argb: '000000' }, // Letra blanca
              name: 'Calibri Light',
              size: 9,
          };
          firstMarkCell.alignment = {
              vertical: 'middle',
              horizontal: 'left',
          };
          firstMarkCell.border = borderStyle; // Aplica bordes

          // Configurar la celda que contiene el valor del Primer Marcaje
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex}`).value = firstTime;

          // Configurar la celda que contiene el valor del Primer Marcaje
          const firstTimeCell = worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex}`);
          firstTimeCell.value = firstTime;
          firstTimeCell.font = {
              name: 'Calibri Light',
              size: 9,
          };
          firstTimeCell.alignment = {
              vertical: 'middle',
              horizontal: 'center',
          };
          firstTimeCell.border = borderStyle;

          // Configurar la celda "Último Marcaje"
          worksheet.getCell(`B${rowIndex + 1}`).value = "Último Marcaje";
          const lastMarkCell = worksheet.getCell(`B${rowIndex + 1}`);
          lastMarkCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFF' }, // Fondo azul oscuro
          };
          lastMarkCell.font = {
              color: { argb: '000000' }, // Letra blanca
              name: 'Calibri Light',
              size: 9,
          };
          lastMarkCell.alignment = {
              vertical: 'middle',
              horizontal: 'left',
          };
          lastMarkCell.border = borderStyle; // Aplica bordes

          // Configurar la celda que contiene el valor del Último Marcaje
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex + 1}`).value = lastTime;

          // Configurar la celda que contiene el valor del Último Marcaje
          const lastTimeCell = worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex + 1}`);
          lastTimeCell.value = lastTime;
          lastTimeCell.font = {
              name: 'Calibri Light',
              size: 9,
          };
          lastTimeCell.alignment = {
              vertical: 'middle',
              horizontal: 'center', // Centrado
          };
          lastTimeCell.border = borderStyle;

          // Configurar la celda "Horas Reportadas"
          worksheet.getCell(`B${rowIndex + 2}`).value = "Horas Reportadas";
          const reportCell = worksheet.getCell(`B${rowIndex + 2}`);
          reportCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '002060' }, // Fondo azul oscuro
          };
          reportCell.font = {
              color: { argb: 'FFFFFF' }, // Letra blanca
              name: 'Calibri Light',
              size: 9,
              bold: true,
          };
          reportCell.alignment = {
              vertical: 'middle',
              horizontal: 'left',
          };
          reportCell.border = borderStyle; // Aplicar bordes


          // Configurar la celda que contiene las horas reportadas
          const hoursCell = worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex + 2}`);
          hoursCell.value = hoursReported;

          // Aplicar formato especial a las horas reportadas
          if (hoursReported !== "00:00:00" || hoursReported == "00:00:00") {
            hoursCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '8ED973' },
            };
            hoursCell.font = {
              color: { argb: '000000' }, // Letra blanca
              name: 'Calibri Light',
              size: 9,
            };
            hoursCell.alignment = {
                vertical: 'middle',
                horizontal: 'center',
            };
            hoursCell.font = { ...generalFontStyle, bold: true };
          }

          // Aplicar bordes a las celdas
          worksheet.getCell(`B${rowIndex}`).border = borderStyle;
          worksheet.getCell(`B${rowIndex + 1}`).border = borderStyle;
          worksheet.getCell(`B${rowIndex + 2}`).border = borderStyle;
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex}`).border = borderStyle;
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex + 1}`).border = borderStyle;
          hoursCell.border = borderStyle;
        });

        // Aplicar bordes a la fila completa
        applyBordersToRow(worksheet.getRow(rowIndex));
        applyBordersToRow(worksheet.getRow(rowIndex + 1));
        applyBordersToRow(worksheet.getRow(rowIndex + 2));

        rowIndex += 3;
      });
    }

    const fileName = `REPORTE SEMANAL.xlsx`;
    const filePathToSave = path.join(__dirname, fileName);
    await workbook.xlsx.writeFile(filePathToSave);

    res.download(filePathToSave, fileName, err => {
      if (err) throw err;
      fs.unlinkSync(filePath);
      fs.unlinkSync(filePathToSave);
    });
  } catch (error) {
    console.error("Error procesando el reporte semanal:", error);
    res.status(500).send("Error procesando el reporte semanal");
  }
});

//--------------------------------------------------------------------------------------------------------------------------------
// Logica de Reporte Fin de Semana

// Función para normalizar nombres
function normalizeName(name) {
  return name ? name.trim().toUpperCase() : null;
}

// Función para calcular la diferencia de tiempo en formato HH:MM:SS
function calculateHoursDifference(start, end) {
  if (!start || !end) return "00:00:00";
  const msDifference = end - start;
  const hours = Math.floor(msDifference / (1000 * 60 * 60));
  const minutes = Math.floor((msDifference % (1000 * 60 * 60)) / (1000 * 60));
  const seconds = Math.floor((msDifference % (1000 * 60)) / 1000);
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

// Función para formatear el tiempo en HH:MM:SS
function formatTime(date) {
  if (!date || isNaN(date)) return "00:00:00";
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${hours}:${minutes}:${seconds}`;
}

// Función para verificar si una fecha es válida y pertenece a días laborables
function isWeekendday(date) {
  const day = date.getDay();
  return day === 6 || day === 0; // Sabado = 6, Domingo = 0
}

// Función para obtener el nombre del día de la semana (solo sábado y domingo)
function getDayNameWeekend(date) {
  const dayNames = ["DOMINGO", "SÁBADO"];
  const dayIndex = date.getDay(); // Obtiene el índice del día directamente
  if (dayIndex === 6) return `${dayNames[1]} ${String(date.getDate()).padStart(2, '0')}`; // Sábado
  if (dayIndex === 0) return `${dayNames[0]} ${String(date.getDate()).padStart(2, '0')}`; // Domingo
  return null; // Si no es sábado ni domingo, retorna null
}

// Función para actualizar el primer y último marcaje
function updateEntry(entry, entryTime) {
  if (!entry.first || new Date(entryTime) < new Date(entry.first)) {
    entry.first = entryTime;
  }
  if (!entry.last || new Date(entryTime) > new Date(entry.last)) {
    entry.last = entryTime;
  }
}

// Función para aplicar bordes a todas las celdas
function applyBordersToRow(row) {
  row.eachCell(cell => {
    cell.border = borderStyle;
  });
}

// Endpoint para procesar y generar el reporte del fin de semana
app.post('/upload_weekend', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    let fileContent = fs.readFileSync(filePath, 'utf-8');

    // Filtrar encabezados innecesarios
    const lines = fileContent.split('\n').filter(line => line.trim() !== '');
    if (lines[0].includes("Todos los Eventos")) {
      lines.shift();
    }
    fileContent = lines.join('\n');

    // Leer y parsear los registros del CSV
    const records = parse(fileContent, {
      columns: true,
      bom: true,
      skip_empty_lines: true,
      relax_column_count: true,
    });

    // Filtrar registros solo de lunes a viernes y con fechas válidas
    const filteredRecords = records.filter(record => {
      const entryTime = new Date(record.Tiempo);
      return isWeekendday(entryTime) && !isNaN(entryTime);
    });

    // Extraer fechas únicas y ordenarlas
    const uniqueDates = [...new Set(filteredRecords.map(record => {
      const date = new Date(record.Tiempo);
      return date.toISOString().split('T')[0];
    }))].sort();

    // Convertir fechas únicas a formato legible
    const daysOfWeek = uniqueDates.map(dateStr => {
      const date = new Date(dateStr);
      return getDayNameWeekend(date);
    }).filter(Boolean); // Filtrar valores nulos o inválidos

    const workbook = new ExcelJS.Workbook();

    // Generar una hoja para cada grupo de usuarios
    for (const [sheetName, users] of Object.entries(userGroups)) {
      const worksheet = workbook.addWorksheet(sheetName);

      // Configurar columnas
      worksheet.columns = [
        { header: "NOMBRE", key: "name", width: 30 },
        { header: "REPORTE", key: "report", width: 20 },
        ...daysOfWeek.map(day => ({ header: day, key: day, width: 20 })),
      ];

      // Estilo de encabezados
      worksheet.getRow(1).eachCell(cell => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '002060' },
        };
        cell.font = { ...generalFontStyle, color: { argb: 'FFFFFF' }, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = borderStyle;
      });

      // Inicializar datos de usuarios
      const userEntries = {};
      users.forEach(user => {
        userEntries[user] = uniqueDates.map(() => ({ first: null, last: null }));
      });

      // Procesar registros
      filteredRecords.forEach(record => {
        const fullName = normalizeName(`${record.Nombre} ${record.Apellido}`);
        const entryTime = new Date(record.Tiempo);
        if (!users.map(normalizeName).includes(fullName) || isNaN(entryTime)) return;

        const entryDate = entryTime.toISOString().split('T')[0];
        const dayIndex = uniqueDates.indexOf(entryDate);
        if (dayIndex !== -1) {
          const entry = userEntries[fullName][dayIndex];
          updateEntry(entry, entryTime);
        }
      });

      // Rellenar datos en la hoja
      let rowIndex = 2;
      users.forEach(user => {
        const displayName = NAME_MAPPING[user] || user; // Usa el nombre mapeado si existe, o el original
        const entries = userEntries[user];
        worksheet.mergeCells(`A${rowIndex}:A${rowIndex + 2}`);
        const nameCell = worksheet.getCell(`A${rowIndex}`);
        nameCell.value = displayName; // Escribe el nombre mapeado
        nameCell.alignment = centeredAlignment;
        nameCell.font = { ...generalFontStyle, bold: true };
        nameCell.border = borderStyle;

        uniqueDates.forEach((date, index) => {
          const entry = entries[index];
          const firstTime = entry.first ? formatTime(entry.first) : "00:00:00";
          const lastTime = entry.last ? formatTime(entry.last) : "00:00:00";
          const hoursReported = entry.first && entry.last
            ? calculateHoursDifference(entry.first, entry.last)
            : "00:00:00";

          // Rellenar celdas con los datos
          worksheet.getCell(`B${rowIndex}`).value = "Primer Marcaje";
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex}`).value = firstTime;

          worksheet.getCell(`B${rowIndex + 1}`).value = "Último Marcaje";
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex + 1}`).value = lastTime;

          worksheet.getCell(`B${rowIndex + 2}`).value = "Horas Reportadas";
          const hoursCell = worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex + 2}`);
          hoursCell.value = hoursReported;

          // Aplicar formato especial a las horas reportadas
          if (hoursReported !== "00:00:00" || hoursReported == "00:00:00") {
            hoursCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: '8ED973' },
            };
            hoursCell.font = { ...generalFontStyle, bold: true };
          }

          // Aplicar bordes a las celdas
          worksheet.getCell(`B${rowIndex}`).border = borderStyle;
          worksheet.getCell(`B${rowIndex + 1}`).border = borderStyle;
          worksheet.getCell(`B${rowIndex + 2}`).border = borderStyle;
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex}`).border = borderStyle;
          worksheet.getCell(`${String.fromCharCode(67 + index)}${rowIndex + 1}`).border = borderStyle;
          hoursCell.border = borderStyle;
        });

        // Aplicar bordes a la fila completa
        applyBordersToRow(worksheet.getRow(rowIndex));
        applyBordersToRow(worksheet.getRow(rowIndex + 1));
        applyBordersToRow(worksheet.getRow(rowIndex + 2));

        rowIndex += 3;
      });
    }

    const fileName = `REPORTE FIN DE SEMANA.xlsx`;
    const filePathToSave = path.join(__dirname, fileName);
    await workbook.xlsx.writeFile(filePathToSave);

    res.download(filePathToSave, fileName, err => {
      if (err) throw err;
      fs.unlinkSync(filePath);
      fs.unlinkSync(filePathToSave);
    });
  } catch (error) {
    console.error("Error procesando el reporte de fin de semana:", error);
    res.status(500).send("Error procesando el reporte de fin de semana");
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Servidor corriendo en el puerto ${PORT}`));