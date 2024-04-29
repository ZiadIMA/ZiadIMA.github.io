Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("button").onclick = addUrgency;
  }
});


function addUrgency() {
  const item = Office.context.mailbox.item;
  const select = document.getElementById("urgencySelect");
  const urgencyValue = select.value;
  const urgencyText = select.options[select.selectedIndex].text;
  const machineName = document.getElementById("machineName").value; // Récupérer le nom de la machine
  const lineName = document.getElementById("lineName").value; // Récupérer le nom de la ligne

  // Utiliser getAsync pour récupérer le sujet de l'élément
  item.subject.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      write(asyncResult.error.message);
      return;
    }

    let subject = asyncResult.value;
    const prefixRegex = /\[[^\]]+\]/; // Expression régulière pour trouver le préfixe existant

    // Vérifier si le sujet contient déjà un préfixe
    if (prefixRegex.test(subject)) {
      // Remplacer le préfixe existant par le nouveau
      subject = subject.replace(prefixRegex, `[${lineName}/${machineName}/@${urgencyValue}]`);
    } else {
      // Ajouter le préfixe au début du sujet
      subject = `[${lineName}/${machineName}/@${urgencyValue}] - ${subject}`;
    }

    // Mettre à jour le sujet avec la nouvelle valeur
    item.subject.setAsync(subject, { coercionType: Office.CoercionType.Text });
  });
}
