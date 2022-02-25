const readXlsxFile = require('read-excel-file/node');
const writeXlsxFile = require('write-excel-file/node')

const XLSX_INPUT = '/Users/martip/Developer/Exentriq/hr/MOTTOLA_122021.xlsx';
const XLSX_OUTPUT = '/Users/martip/Developer/Exentriq/hr/2021_12.xlsx';

const ITALIAN_HOLIDAYS = [
  { day:  1, month:  1 }, // Capodanno
  { day:  6, month:  1 }, // Epifania
  { day: 25, month:  4 }, // Festa della Liberazione
  { day:  1, month:  5 }, // Festa dei lavoratori
  { day:  2, month:  6 }, // Festa della Repubblica Italiana
  { day: 15, month:  8 }, // Ferrragosto
  { day:  1, month: 11 }, // Ognissanti
  { day:  8, month: 12 }, // Immacolata Concezione
  { day: 25, month: 12 }, // Natale
  { day: 26, month: 12 }, // Santo Stefano
];

const SPECIAL_DAYS = [
  { day:  7, month: 12, hours: 0 }, // S. Ambrogio
  { day: 24, month: 12, hours: 4 }, // Vigilia di Natale
];

const getEasterMonday = (year) => {
  const c = Math.floor(year / 100);
  const n = year - 19 * Math.floor(year / 19);
  const k = Math.floor((c - 17) / 25);
  let i = c - Math.floor(c / 4) - Math.floor((c - k) / 3) + 19 * n + 15;
  i = i - 30 * Math.floor((i / 30));
  i = i - Math.floor(i / 28) * (1 - Math.floor(i / 28) * Math.floor(29 / (i + 1)) * Math.floor((21 - n) / 11));
  let j = year + Math.floor(year / 4) + i + 2 - c + Math.floor(c / 4);
  j = j - 7 * Math.floor(j / 7);
  const l = i - j;
  const m = 3 + Math.floor((l + 40) / 44);
  const d = l + 28 - 31 * Math.floor(m / 4);
  let easterMondayDate = new Date(year, m - 1, d);
  easterMondayDate.setDate(easterMondayDate.getDate() + 1);
  easterMondayDate.setMinutes(easterMondayDate.getMinutes() - easterMondayDate.getTimezoneOffset());
  return easterMondayDate;
};

const isHoliday = (year, month, day) => {
  if (ITALIAN_HOLIDAYS.find(x => x.month === month && x.day === day)) {
    return true;
  }
  return false;
};

const getDayWorkingHours = (month, day, weekDay, reduced) => {
  const specialDay = SPECIAL_DAYS.find(x => x.month === month && x.day === day);
  return specialDay ? specialDay.hours : (weekDay === 5 && reduced) ? 7 : 8;
};

const calculateWorkingHours = (year, month, reduced) => {

  const weekday = [ 'domenica', 'lunedì', 'martedì', 'mercoledì', 'giovedì', 'venerdì', 'sabato' ];
  const easterMonday = getEasterMonday(year);
  
  let workingHours = 0;

  for (let day = 1; day <= new Date(year, month, 0).getDate(); day++) {
    const date = new Date(year, month - 1, day);
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    const weekDay = date.getDay();
    if (
      weekDay > 0 && weekDay < 6
      && !isHoliday(year, month, day)
      && date.toISOString() !== easterMonday.toISOString()
    ) {
      workingHours += getDayWorkingHours(month, day, weekDay, reduced);
      // console.log(`${day}/${month}/${year}: ${weekday[weekDay]}`);
    }
  }
  return workingHours;
};

const parseMonth = (monthString) => {

  const result = { month: null, year: null };
  const months = [
    'gennaio', 'febbraio', 'marzo', 'aprile',
    'maggio', 'giugno', 'luglio', 'agosto',
    'settembre', 'ottobre', 'novembre', 'dicembre'
  ];
  const month = months.findIndex(x => x === monthString.replace(/\d/g, '').trim().toLowerCase());
  if (month > -1) {
    result.month = month + 1;
  }
  const year = parseInt(monthString.replace(/[^\d\s]/g, ''));
  if (!isNaN(year)) {
    result.year = year > 2000 ? year : year + 2000;
  }
  return result;
}

const parseCell = (cell, schema) => {

  let value;
  if (schema.regex) {
    const match = cell.match(schema.regex);
    if (match) {
      value = match[1];
    }
  } else {
    value = cell;
  }

  if (schema.transform) {
    value = schema.transform(value);
  }
  return value;

};

const parseRow = (cells, schema) => {
  let value;
  if (schema.formula) {
    switch (schema.formula) {
      case 'sum':
        value = cells.reduce((acc, obj) => {
          const num = parseInt(obj);
          acc += isNaN(num) ? 0 : num;
          return acc;
        }, 0)
        break;
      default:
        break;
    }
  }
  return value;
};

const parseTimeSheet = (rows) => {

  const schema = {
    name: { row: 0, column: 13, regex: /(?:NOME\s*COGNOME\s*)*(?<name>[\w\s]*)/ },
    period: { row: 0, column: 26, regex: /(?:MESE\s*)*(?<month>[\w\s]*)/, transform: parseMonth },
    work: { row: 2, formula: 'sum', parent: 'hours' },
    vacation: { row: 4, formula: 'sum', parent: 'hours' },
    leave: { row: 6, formula: 'sum', parent: 'hours' },
    paid: { row: 8, formula: 'sum', parent: 'hours' }
  };

  const data = {};

  for (const key of Object.keys(schema)) {
    const schemaItem = schema[key];

    let value;

    if (schemaItem.column) {
      value = parseCell(rows[schemaItem.row][schemaItem.column], schemaItem);
    } else {
      value = parseRow(rows[schemaItem.row], schemaItem);
    }
    if (schemaItem.parent) {
      if (!data[schemaItem.parent]) {
        data[schemaItem.parent] = {};
      }
      data[schemaItem.parent][key] = value;
    } else {
      data[key] = value;
    }
  }

 return data;

};

const parseExpenses = (rows) => {
  const total = parseFloat(rows[44][3]);
  if (!isNaN(total)) {
    return total;
  }
  return 0;
};

(async () => {

  const employees = [];

  const test = await readXlsxFile(XLSX_OUTPUT, { sheet: 1 });
  process.exit(0);

  const timeSheetRows = await readXlsxFile(XLSX_INPUT, { sheet: 1 });
  const employee = parseTimeSheet(timeSheetRows);
  const expensesSheetRows = await readXlsxFile(XLSX_INPUT, { sheet: 2 });
  employee.expenses = parseExpenses(expensesSheetRows);
  employee.anomalies = [];

  const reduced = true; // true -> 39 hours/week, false -> 40 hours/week

  const workingHours = calculateWorkingHours(employee.period.year, employee.period.month, reduced);
  const totalHours = Object.values(employee.hours).reduce((acc, obj) => { return acc + obj; }, 0);

  if (totalHours < workingHours) {
    employee.anomalies.push(`Total hours (${totalHours}) < working hours (${workingHours})!`);
  } else if (totalHours > workingHours) {
    employee.anomalies.push(`Total hours (${totalHours}) > working hours (${workingHours})!`);
  }

  console.log(employee);

  employees.push(employee);

  await writeXlsxFile(employees, {
    schema: [
      { column: 'Dipendente', type: String, width: 30, height: 20, value: employee => employee.name },
      { column: 'Diarie (ore)', type: Number, width: 15, align: 'right', height: 20, value: employee => employee.hours.work },
      { column: 'Ferie (ore)', type: Number, width: 15, align: 'right', height: 20, value: employee => employee.hours.vacation },
      { column: 'Permessi non retribuiti (ore)', type: Number, width: 25, align: 'right', height: 20, value: employee => employee.hours.leave },
      { column: 'Permessi retribuiti (ore)', type: Number, width: 25, align: 'right', height: 20, value: employee => employee.hours.paid },
      { column: 'Spese (€)', type: Number, format: '#,##0.00', width: 10, align: 'right', height: 20, value: employee => employee.expenses },
      { column: 'Anomalie', type: String, width: 50, height: 20, wrap: true, value: employee => employee.anomalies.join('\n') }
    ],
    headerStyle: {
      backgroundColor: '#eeeeee',
      fontWeight: 'bold',
      height: 20
    },
    fontFamily: 'Arial',
    fontSize: 10,
    sheet: 'Dipendenti',
    stickyRowsCount: 1,
    filePath: XLSX_OUTPUT
  });

})();