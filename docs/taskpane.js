const SCRIPT_ID = "DEINE_SCRIPT_ID";
let accessToken = null;

function login() {
  return new Promise((resolve, reject) => {
    const state = crypto.randomUUID();
    const url = `${location.origin}/dialog.html?state=${state}`;
    Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(res.error);
      const dlg = res.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        dlg.close();
        const m = JSON.parse(arg.message);
        if (m.error || m.state !== state) return reject(new Error(m.error || "State stimmt nicht"));
        accessToken = m.token;
        resolve();
      });
    });
  });
}

async function callScript(fn, parameters = []) {
  const r = await fetch(`https://script.googleapis.com/v1/scripts/${SCRIPT_ID}:run`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ function: fn, parameters })
  });
  if (r.status === 401) { accessToken = null; throw new Error("Sitzung abgelaufen, bitte neu anmelden"); }
  const data = await r.json();
  if (data.error) throw new Error(JSON.stringify(data.error.details ?? data.error));
  return data.response.result;
}

async function writeTwoToActiveCell() {
  if (!accessToken) await login();
  const result = await callScript("getTwo");
  await Excel.run(async (ctx) => {
    ctx.workbook.getActiveCell().values = [[result]];
    await ctx.sync();
  });
}

Office.onReady(() => {
  document.getElementById("btn").onclick = () =>
    writeTwoToActiveCell().catch((e) => console.error(e) || alert(e.message));
});
