// Läuft in der Seitenleiste (Task Pane).
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const button = document.getElementById("writeTwo");
    button.disabled = false;
    button.addEventListener("click", writeTwo);
  }
});

async function writeTwo() {
  const status = document.getElementById("status");
  status.textContent = "";

  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell(); // aktuell ausgewählte Zelle
      cell.load("address");
      cell.values = [[2]];
      await context.sync();
      status.textContent = `Eine 2 wurde in ${cell.address} geschrieben.`;
    });
  } catch (error) {
    status.textContent = `Fehler: ${error.message}`;
    console.error(error);
  }
}
