Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      document.getElementById("button").onclick = addUrgency;
  }
});

function addUrgency() {
  const item = Office.context.mailbox.item;
  const select = document.getElementById("urgencySelect");
  const urgencyValue = select.value;
  const urgencyText = select.options[select.selectedIndex].text;
  const machineName = document.getElementById("machineName").value;
  const lineName = document.getElementById("lineName").value;

  item.subject.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(asyncResult.error.message);
      return;
    }

    let subject = asyncResult.value;
    const prefixRegex = /\[[^\]]+\]/;

    if (prefixRegex.test(subject)) {
      subject = subject.replace(prefixRegex, `[${lineName}/${machineName}/@${urgencyValue}]`);
    } else {
      subject = `[${lineName}/${machineName}/@${urgencyValue}] - ${subject}`;
    }

    item.subject.setAsync(subject, { coercionType: Office.CoercionType.Text });
  });
}
