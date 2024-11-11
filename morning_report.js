//morning_report.js

const ExcelJS = require('exceljs');

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

function generateMorningReport(records) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Morning Report");
  
    // Obtener la fecha del primer registro para el encabezado
    const firstDate = new Date(records[0].Tiempo);
    const formattedDate = `${String(firstDate.getDate()).padStart(2, '0')}_${String(firstDate.getMonth() + 1).padStart(2, '0')}_${firstDate.getFullYear()}`;
  
    // Configurar columnas
    worksheet.columns = [
      { header: "Nombre", key: "name", width: 30 },
      { header: "Reporte", key: "report", width: 20 },
      { header: formattedDate, key: "first_entry", width: 20 },
    ];
  
    // Estilos del encabezado
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
  
    // Agregar datos de los empleados en orden de PRIORITY_USERS
    PRIORITY_USERS.forEach(name => {
      const userRecords = records.filter(record => `${record.Nombre} ${record.Apellido}` === name);
      const firstEntry = userRecords.length > 0 ? new Date(userRecords[0].Tiempo) : null;
  
      const row = worksheet.addRow({
        name: name,
        report: "Primer Marcaje",
        first_entry: firstEntry ? `${String(firstEntry.getHours()).padStart(2, '0')}:${String(firstEntry.getMinutes()).padStart(2, '0')}:${String(firstEntry.getSeconds()).padStart(2, '0')}` : "",
      });
  
      // Estilos para cada fila
      row.eachCell((cell, colNumber) => {
        cell.font = { name: 'Calibri Light', size: 9 };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        if (colNumber === 3) { // Columna de "Primer Marcaje"
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      });
    });
  
    return workbook;
  }
  
  module.exports = generateMorningReport;
