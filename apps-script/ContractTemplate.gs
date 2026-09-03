// Bundet script på kontraktmalen "1 . Kontrakt - BHTN - MALFIL AS".
// Kopieres automatisk med filen når DriveKunder.gs lager en ny kundekontrakt
// via makeCopy() — så BHTN-menyen dukker opp på alle NYE kundekopier.
//
// MALFIL_DOC_ID er ID-en til selve malfilen (samme verdi som SAMPLE_CONTRACT_ID
// i DriveKunder.gs). En kopi laget med makeCopy() får alltid en annen ID, så
// denne sjekken kan aldri gi falskt utslag — og krever ingen ny Drive-tilgang.
const MALFIL_DOC_ID = "1RFq0e1X61u-6IxuMYX-t8wOf5-z-iu6YtOOmT3cmWsI";

function onOpen() {
  DocumentApp.getUi()
    .createMenu("BHTN")
    .addItem("Fjern gul markering", "clearYellowHighlight")
    .addToUi();
}

function clearYellowHighlight() {
  const doc = DocumentApp.getActiveDocument();

  if (doc.getId() === MALFIL_DOC_ID) {
    DocumentApp.getUi().alert(
      "Dette er malfilen — den skal ikke endres. Kjør dette kun på en kundekopi."
    );
    return;
  }

  const body = doc.getBody();
  const text = body.editAsText();
  const len = text.getText().length;
  if (len > 0) {
    text.setBackgroundColor(0, len - 1, null);
  }

  DocumentApp.getUi().alert("Gul markering fjernet.");
}
