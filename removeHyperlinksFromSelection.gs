function onOpen() {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Remove all hyperlinks', 'init')
    .addToUi()
}

function getAllLinks(element, selectionStartOffset, selectionEndOffsetInclusive) {
  var links = []
  var type = element.getType()
  if (type === DocumentApp.ElementType.TEXT) {
    var textObj = element.editAsText()
    var text = element.getText()
    var inUrl = false
    
    var firstStartOffset = 0
    if (selectionStartOffset !== -1) {
      firstStartOffset = selectionStartOffset
    }
    
    var lastEndOffsetInclusive = text.length
    if (selectionEndOffsetInclusive !== -1) {
      lastEndOffsetInclusive = selectionEndOffsetInclusive + 1
    }
    
    for (var ch = firstStartOffset; ch < lastEndOffsetInclusive; ch++) {
      var url = textObj.getLinkUrl(ch)
      if (url != null) {
        if (!inUrl) {
          // We are now!
          inUrl = true
          var curUrl = {}
          curUrl.element = element
          curUrl.url = url
          curUrl.startOffset = ch
        } else {
          curUrl.endOffsetInclusive = ch
          if (ch === text.length - 1) {
            // this hyperlink is the end of the TEXT element
            inUrl = false
            links.push(curUrl)
            curUrl = {}
          }
        }          
      } else {
        if (inUrl) {
          // Not any more, we're not.
          inUrl = false
          links.push(curUrl)  // add to links
          curUrl = {}
        }
      }
    }
  }
  
  var singletonElement
  if (type === DocumentApp.ElementType.INLINE_IMAGE) {
    singletonElement = element.asInlineImage()
  } else if (type === DocumentApp.ElementType.EQUATION) {
    singletonElement = element.asEquation()
  } else if (type === DocumentApp.ElementType.EQUATION_FUNCTION) {
    singletonElement = element.asEquationFunction()
  } else if (type === DocumentApp.ElementType.LIST_ITEM) {
    singletonElement = element.asListItem()
  } else if (type === DocumentApp.ElementType.PARAGRAPH) {
    singletonElement = element.asParagraph()
  } else if (type === DocumentApp.ElementType.TABLE) {
    singletonElement = element.asTable()
  } else if (type === DocumentApp.ElementType.TABLE_CELL) {
    singletonElement = element.asTableCell()
  } else if (type === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
    singletonElement = element.asTableOfContents()
  } else if (type === DocumentApp.ElementType.TABLE_ROW) {
    singletonElement = element.asTableRow()
  } else if (type === DocumentApp.ElementType.TEXT) {
    singletonElement = element.asText()
  }
  
  if (singletonElement) {
    var link = singletonElement.getLinkUrl()
    if (link) {
      links.push({
        element: singletonElement,
        url: link,
      })
    }
  }
  if (element.getNumChildren) {
    var numChildren = element.getNumChildren()
    for (var i = 0; i < numChildren; i++) {
      links = links.concat(getAllLinks(element.getChild(i), selectionStartOffset, selectionEndOffsetInclusive))
    }
  }
  return links
}
function init() {
  var doc = DocumentApp.getActiveDocument()
  var selection = doc.getSelection()
  if (selection) {
    var elements = selection.getRangeElements()
    for (var i = 0; i < elements.length; i++) {
      var rangeElement = elements[i]
      var element = rangeElement.getElement()
      
      if (element.editAsText) {
        var text = element.editAsText()
        
        var links = getAllLinks(element, rangeElement.getStartOffset(), rangeElement.getEndOffsetInclusive())
        for (var j = 0; j < links.length; j++) {
          var link = links[j]
          if (link.element.setLinkUrl) {
            if (link.hasOwnProperty('startOffset')) {
              link.element.setLinkUrl(link.startOffset, link.endOffsetInclusive, '')
            } else {
              link.element.setLinkUrl(null)
            }
          }
        }
      }
      
    }
  } else {
    DocumentApp.getUi().alert('Select the text with hyperlinks to be removed')
  }
}
