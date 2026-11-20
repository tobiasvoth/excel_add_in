// Wird beim Klick auf den Button im Menüband ausgeführt.
Office.onReady(() => {
  // Excel ist bereit
});

async function writeTwo(event) {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell(); // aktuell ausgewählte Zelle
      cell.values = [[2]];
      await context.sync();
    });
  } catch (error) {
    console.error("Fehler beim Schreiben:", error);
  }

  // Wichtig: Excel mitteilen, dass die Funktion fertig ist
  event.completed();
}

// Funktionsname muss mit <FunctionName> im Manifest übereinstimmen
Office.actions.associate("writeTwo", writeTwo);
