function onOpen() {
  const menu = SlidesApp.getUi().createMenu('Custom Menu')
  menu.addItem('Make Title', 'this.myTitle')
  menu.addItem('Make Text', 'this.myText')
  menu.addItem('Text Except Red', 'this.myColorChange')
  menu.addItem('Notes, Blue', 'this.myNotesBlue')
  menu.addItem('Notes, Black', 'this.myNotesBlack')
  menu.addItem('FG to presenter notes', 'this.myFG')
  menu.addToUi()
}

function normify(elem) {
  elem.setBold(false)
  elem.setStrikethrough(false)
  elem.setUnderline(false)
  elem.setForegroundColor('#77787b')
}

function titleStyle(elem) {
  normify(elem)
  elem.setFontFamily('Copse')
  elem.setFontSize(28)
}

function textStyle(elem) {
  normify(elem)
  elem.setFontFamily('Open Sans')
  elem.setFontSize(20)
}

function notesHelper(bool) {
  const sel = SlidesApp.getActivePresentation().getSelection().getTextRange().getTextStyle()
  if (bool) {
    sel.setForegroundColor('#0432ff')
  }
  else {
    sel.setForegroundColor('#000000')
  }
  sel.setFontSize(11)
  sel.setFontFamily('Open Sans')
  myFG()
}

function styleHelper(bool) {
  const sel = SlidesApp.getActivePresentation().getSelection()
  if (sel.getSelectionType() == SlidesApp.SelectionType.TEXT) {
    elemSelect = sel.getTextRange().getTextStyle()
    if (bool) {
      this.titleStyle(elemSelect)
    }
    else {
      this.textStyle(elemSelect)
    }
  }
  else {
    const elemArray = sel.getPageElementRange().getPageElements()
    for (let i = 0; i < elemArray.length; i++) {
      try {
        let elem = elemArray[i].asShape().getText().getTextStyle()
        if (bool) {
          this.titleStyle(elem)
        }
        else {
          this.textStyle(elem)
        }
      }
      catch {
        let tab = elemArray[i].asTable()
        for (let j = 0; j < tab.getNumColumns(); j++) {
          for (let k = 0; k < tab.getNumRows(); k++) {
            let cell = tab.getCell(k, j).getText().getTextStyle()
            if (bool) {
              this.titleStyle(cell)
            }
            else {
              this.textStyle(cell)
            }
          }
        }
      }
    }
  }
}

function colorChangeHelper() {
  const text = SlidesApp.getActivePresentation().getSelection().getTextRange()
  var start = text.getStartIndex()
  var end = text.getEndIndex()
  for (let i = start; i < end; i++) {
    try {
      if (text.getRange(i - start, i - start + 1).getTextStyle().getForegroundColor().asRgbColor().getRed() < 190) {
        text.getRange(i - start, i - start + 1).getTextStyle().setForegroundColor('#77787b')
      }
      else {
        text.getRange(i - start, i - start + 1).getTextStyle().setForegroundColor('#dc4a05')
      }
    }
    catch {
      if (text.getRange(i - start, i - start + 1).getTextStyle().getForegroundColor().asThemeColor().getThemeColorType() == 'DARK1') {
        text.getRange(i - start, i - start + 1).getTextStyle().setForegroundColor('#595959')
      }
      else {
        text.getRange(i - start, i - start + 1).getTextStyle().setForegroundColor('#dc4a05')
      }
    }
  }
}

function myTitle() {
  styleHelper(true)
}

function myText() {
  styleHelper(false)
}

function myColorChange() {
  colorChangeHelper()
}

function myNotesBlue() {
  notesHelper(true)
}

function myNotesBlack() {
  notesHelper(false)
}

function myFG() {
  const sel = SlidesApp.getActivePresentation().getSelection().getTextRange()
  sel.replaceAllText('● ', '\n\n')
  sel.replaceAllText('○ ', '\n\n')
  sel.replaceAllText('\n\n', '\n*')
  sel.replaceAllText('\n', ' ')
  sel.replaceAllText('*', '\n\n')
  sel.replaceAllText('\n\n ', '\n\n')
}