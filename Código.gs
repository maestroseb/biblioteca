/**
 * Sirve la página web principal al usuario.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle(SpreadsheetApp.getActiveSpreadsheet().getName())
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Obtiene los datos iniciales de las hojas de cálculo.
 */
function getInitialData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = spreadsheet.getSheetByName("BaseDeDatos");
  const loansSheet = spreadsheet.getSheetByName("Prestamos");
  
  const singleLoanRuleActive = dbSheet.getRange("D1").getValue() === true;
  
  // MEJORA: Acotar rangos con getLastRow() para no leer miles de filas vacías.
  const dbLastRow = dbSheet.getLastRow();
  const allStudents = dbLastRow >= 2
    ? dbSheet.getRange("A2:A" + dbLastRow).getValues().flat().map(s => String(s).trim()).filter(Boolean)
    : [];
  const allBooks = dbLastRow >= 2
    ? dbSheet.getRange("B2:B" + dbLastRow).getValues().flat().map(s => String(s).trim()).filter(Boolean)
    : [];
  
  if (loansSheet.getLastRow() < 2) {
    return { 
      spreadsheetName: spreadsheet.getName(),
      singleLoanRuleActive,
      students: allStudents, 
      availableBooks: allBooks, 
      activeLoans: [] 
    };
  }
  
  const loansDataRange = loansSheet.getRange("A2:E" + loansSheet.getLastRow());
  const loansData = loansDataRange.getValues();

  const allLoansWithRowNumbers = loansData.map((row, index) => ({
    rowNumber: index + 2, 
    student: String(row[1]).trim(), 
    book: String(row[2]).trim(),
    loanDate: (row[3] instanceof Date) ? row[3].toISOString() : null,
    returnDate: row[4]
  }));

  const activeLoans = allLoansWithRowNumbers.filter(loan => loan.returnDate === "" || loan.returnDate == null);
  const loanedBooksTitles = activeLoans.map(loan => loan.book);
  const availableBooks = allBooks.filter(book => !loanedBooksTitles.includes(book));

  let availableStudents;
  if (singleLoanRuleActive) {
    const studentsWithLoans = new Set(activeLoans.map(loan => loan.student));
    availableStudents = allStudents.filter(student => !studentsWithLoans.has(student));
  } else {
    availableStudents = allStudents;
  }
    
  return { 
    spreadsheetName: spreadsheet.getName(),
    singleLoanRuleActive,
    students: availableStudents,
    availableBooks, 
    activeLoans 
  };
}

/**
 * Procesa la hoja de préstamos y devuelve un objeto con estadísticas.
 */
function getStatistics() {
  const loansSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prestamos");
  if (loansSheet.getLastRow() < 2) {
    return { topBooks: [], topReaders: [], simpleStats: { totalLoans: 0, uniqueBooks: 0, uniqueReaders: 0 } };
  }

  const loansData = loansSheet.getRange("B2:C" + loansSheet.getLastRow()).getValues();
  const bookCounts = {};
  const readerCounts = {};
  
  loansData.forEach(row => {
    const reader = String(row[0]).trim(); 
    const book = String(row[1]).trim();
    if (reader) readerCounts[reader] = (readerCounts[reader] || 0) + 1;
    if (book) bookCounts[book] = (bookCounts[book] || 0) + 1;
  });

  const sortAndSlice = (counts) => Object.entries(counts)
      .sort(([, a], [, b]) => b - a).slice(0, 10)
      .map(([name, count]) => ({ name, count }));

  return {
    topBooks: sortAndSlice(bookCounts),
    topReaders: sortAndSlice(readerCounts),
    simpleStats: {
      totalLoans: loansData.filter(r => String(r[0]).trim() || String(r[1]).trim()).length,
      uniqueBooks: Object.keys(bookCounts).length,
      uniqueReaders: Object.keys(readerCounts).length
    }
  };
}

/**
 * Registra un préstamo con protección contra concurrencia.
 */
function loanBook(student, book) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Espera hasta 10 segundos
  } catch (e) {
    throw new Error("El sistema está ocupado, inténtalo de nuevo en unos segundos.");
  }

  try {
    const loansSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prestamos");
    const lastRow = loansSheet.getLastRow();
    const nextRow = lastRow + 1;
    const lastLoanNumber = lastRow >= 2 ? (Number(loansSheet.getRange(lastRow, 1).getValue()) || 0) : 0;
    const newLoanNumber = lastLoanNumber + 1;
    const loanDate = new Date();
    const newLoanData = [newLoanNumber, student, book, loanDate, ""];
    loansSheet.getRange(nextRow, 1, 1, 5).setValues([newLoanData]);
    SpreadsheetApp.flush(); // Asegura la escritura antes de soltar el lock
    return { rowNumber: nextRow, student, book, loanDate: loanDate.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Registra la devolución de un libro.
 */
function returnBook(loanRow, bookTitle, studentName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("El sistema está ocupado, inténtalo de nuevo en unos segundos.");
  }

  try {
    const loansSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prestamos");
    loansSheet.getRange(loanRow, 5).setValue(new Date());
    SpreadsheetApp.flush();
    return { 
      returnedRow: loanRow, 
      returnedBookTitle: bookTitle,
      returnedStudent: studentName 
    };
  } finally {
    lock.releaseLock();
  }
}