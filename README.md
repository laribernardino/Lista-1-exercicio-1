# Lista-1-exercicio-1
function onlyNumberKey(evt) {
  let ASCIICode = evt.which ? evt.which : evt.keyCode;
  if (ASCIICode > 31 && (ASCIICode < 48 || ASCIICode > 57)) return false;
  return true;
}
function buttonCalc() {
  Word.run(function(context) {
    let docBody = context.document.body;
    let n1 = txtn1.value;
    let n2 = txtn2.value;
    let n3 = txtn3.value;
    let n4 = txtn4.value;
    let n5 = txtn5.value;
    docBody.clear();
    if (n1 == "" || n2 == "" || n3 == "" || n4 == "" || n5 == "") {
      docBody.insertParagraph("Preste atenção. Não deixe os campos em branco.", "End");
    } else {
      let results = [];
      n1 = parseInt(n1);
      n2 = parseInt(n2);
      n3 = parseInt(n3);
      n4 = parseInt(n4);
      n5 = parseInt(n5);
      docBody.insertParagraph(`${n1} ^ 2 = ${Math.pow(n1, 2)}`, "End");
      docBody.insertParagraph(`${n2} ^ 2 = ${Math.pow(n2, 2)}`, "End");
      docBody.insertParagraph(`${n3} ^ 2 = ${Math.pow(n3, 2)}`, "End");
      docBody.insertParagraph(`${n4} ^ 2 = ${Math.pow(n4, 2)}`, "End");
      docBody.insertParagraph(`${n5} ^ 2 = ${Math.pow(n5, 2)}`, "End");
    }
    return context.sync();
  });
}
btnCalc.addEventListener("click", buttonCalc);
