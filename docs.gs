function onOpen() {
  const menu = DocumentApp.getUi().createMenu('Custom Menu')
  menu.addItem('Add 1', 'this.addOne')
  menu.addToUi()
}

//adds 1 to the number in each cell
function addOne(){
  const sel = DocumentApp.getActiveDocument().getSelection().getRangeElements()
  for (let i = 0; i < sel.length; i++){
    let tab = sel[i].getElement().editAsText()
    if (tab.getText() != ''){
      let num = Number(tab.getText())
      let numadded = num + 1
      tab.setText(numadded)
    } 
  }
}
