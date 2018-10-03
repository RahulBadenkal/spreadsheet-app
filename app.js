(function () {
  "use strict"

  angular.module("spreadsheet-app", [])
    .controller("MainCtrl", MainCtrl)
    .service("Service1", Service1)
    .service("Helpers", Helpers);


  MainCtrl.$inject = ["Service1", "$scope", "$interval"];
  function MainCtrl(Service1, $scope, $interval) {
    var ctrl = this;
    var db = {
      'sheet1': {
        'formula':
          [
            ['ledgerName', 'tbDebit', 'tbCredit'],
            ['Capital Accounts', 100000, 0],
            ['Shinchan', 60000, 0],
            ['Nohara', 40000, 0]
          ],

        'metadata':
          [
            [{ 'dtype': 'str', 'value': 'ledgerName', 'formula': 'ledgerName', 'format': {} }, { 'dtype': 'str', 'value': 'tbDebit', 'formula': 'tbDebit', 'format': {} }, { 'dtype': 'str', 'value': 'tbCredit', 'formula': 'tbCredit', 'format': {} }],
            [{ 'dtype': 'str', 'value': 'Capital Accounts', 'formula': 'Capital Accounts', 'format': {} }, { 'dtype': 'float', 'value': 100000, 'formula': '100000', 'format': {} }, { 'dtype': 'float', 'value': 0, 'formula': '0', 'format': {} }],
            [{ 'dtype': 'str', 'value': 'Shinchan', 'formula': 'Shinchan', 'format': {} }, { 'dtype': 'float', 'value': 60000, 'formula': '60000', 'format': {} }, { 'dtype': 'float', 'value': 0, 'formula': '0', 'format': {} }],
            [{ 'dtype': 'str', 'value': 'Nohara', 'formula': 'Nohara', 'format': {} }, { 'dtype': 'float', 'value': 40000, 'formula': '40000', 'format': {} }, { 'dtype': 'float', 'value': 0, 'formula': '0', 'format': {} }]
          ],
      }
    };

    // registering my custom renderer
    Handsontable.renderers.registerRenderer('my.custom', Service1.customRenderer);

    // Getting the worksheet div
    ctrl.sheet1Div = document.getElementById('example');

    ctrl.boldButton = document.getElementById('bold-button');

    // Creating and setting new handsontable instance
    ctrl.hot = new Handsontable(ctrl.sheet1Div, {
      data: db.sheet1.formula,
      minRows: 3000,
      minCols: 20,
      minSpareRows: 1,
      rowHeaders: true,
      colHeaders: true,
      stretchH: 'all',
      // Doesn't loose track of the user selected cells when clicking outside the table dom.
      outsideClickDeselects: false,
      manualColumnMove: true,
      manualRowMove: true,
      manualColumnResize: true,
      manualRowResize: true,
      undo: true,
      contextMenu: true,
      // applying custom renderer (after cell metadata is applied)
      renderer: 'my.custom',
      // For initializing cell metadata
      cell: Service1.getMetadata(db.sheet1.metadata),
    });

    // Input type elements of toolbar
    // Not implemented yet
    ctrl.toolBarInputList = ['fontName', 'fontSize', 'fontColour', 'backgroundColour',]

    // Getting the toolbar button divs
    ctrl.toolbarButtonsList = ['bold', 'italic', 'underline', 'textAlignLeft', 'textAlignCenter', 'textAlignRight', 'verticalAlignTop', 'verticalAlignCenter', 'verticalAlignRight']
    ctrl.toolbarButtonsDiv = {}; // key=ButtonName, value=ButtonElement 
    for (var i = 0; i < ctrl.toolbarButtonsList.length; i++) {
      ctrl.toolbarButtonsDiv[ctrl.toolbarButtonsList[i]] = document.getElementById(ctrl.toolbarButtonsList[i]);
    }

    // Initializing button status
    ctrl.buttonsWithStatusList = ctrl.toolbarButtonsList;
    ctrl.buttonStatusObj = {};
    for (var i = 0; i < ctrl.buttonsWithStatusList.length; i++) {
      ctrl.buttonStatusObj[ctrl.buttonsWithStatusList[i]] = false;
    }

    // Check every button's Status after certain interval
    ctrl.updateButtonStatus = $interval(function () {
      if (ctrl.hot !== undefined || ctrl.hot !== null) {
        for (var i = 0; i < ctrl.buttonsWithStatusList.length; i++) {
          var button = ctrl.buttonsWithStatusList[i];
          ctrl.buttonStatusObj[button] = Service1.isSelectionActive(ctrl.hot, button)
        }
      }
      // console.log(ctrl.buttonStatusObj);
    }, 300);

    $scope.$watch("ctrl.buttonStatusObj", function (oldValue, newValue) {
      for (var key in ctrl.buttonStatusObj) {
        if (ctrl.buttonStatusObj.hasOwnProperty(key)) {
          if (ctrl.buttonStatusObj[key] === true) {
            ctrl.toolbarButtonsDiv[key].className = 'toolbar-button-active';
          }
          else {
            ctrl.toolbarButtonsDiv[key].className = 'toolbar-button-inactive';
          }
        }
      }
    }, true)

    // On button click execute this function
    ctrl.change = function (event) {
      var id = event.target.id;
      Service1.applyFormat(ctrl.hot, id, ctrl.buttonStatusObj[id])
    }


  }

  //---------------------------------------------------------------------------------------------------------------------//
  //------------------------------------------Service1 Starts------------------------------------------------------------//
  //---------------------------------------------------------------------------------------------------------------------//

  // Service1, handsontable related functions
  Service1.$inject = ["Helpers"];
  function Service1(Helpers) {
    var service = this;

    // Converts metadata array into suitable data type required by handsontable
    service.getMetadata = function (metaDataArr) {
      var newArr = []
      var rowCount = metaDataArr.length;
      for (var i = 0; i < rowCount; i++) {
        var colCount = metaDataArr[i].length;
        for (var j = 0; j < colCount; j++) {
          newArr.push({ row: i, col: j, metadata: Helpers.clone(metaDataArr[i][j]) });
        }
      }
      return newArr;
    }

    // My defined renderer 
    service.customRenderer = function (instance, td, row, col, prop, value, cellProperties) {
      // console.log("Here in renderer");
      // console.log(value);
      // console.log(row, col);
      // console.log(cellProperties);

      if (cellProperties.hasOwnProperty('metadata')) {
        var metadata = cellProperties.metadata;
        metadata.formula = value;
        var showValue = value;
        metadata.value = showValue;

        if (td !== undefined && td !== null) {
          var format = metadata.format;
          // Set bold
          if (format.hasOwnProperty('bold') && format.bold === true)
            td.style.fontWeight = 'bold';
          else
            td.style.fontWeight = 'normal';

          // Set italics
          if (format.hasOwnProperty('italic') && format.italic === true)
            td.style.fontStyle = 'italic';
          else
            td.style.fontStyle = 'normal';

          // Set underline
          if (format.hasOwnProperty('underline') && format.underline === true)
            td.style.textDecoration = 'underline';
          else
            td.style.textDecoration = 'none';

          // Set Horizontal align 
          if (format.hasOwnProperty('textAlign')) {
            if (format.textAlign === 'textAlignLeft')
              td.style.textAlign = 'left';
            else if (format.textAlign === 'textAlignCenter')
              td.style.textAlign = 'center';
            else if (format.textAlign === 'textAlignRight')
              td.style.textAlign = 'right';
          }
          else
            td.style.textAlign = 'left'

          // Set Vertical align 
          if (format.hasOwnProperty('verticalAlign')) {
            if (format.verticalAlign === 'verticalAlignTop'){
              // debugger;
              td.style.verticalAlign = 'top';
            }
            else if (format.verticalAlign === 'verticalAlignCenter')
              td.style.verticalAlign = 'middle';
            else if (format.verticalAlign === 'verticalAlignRight')
              td.style.verticalAlign = 'bottom';
          }
          else
            td.style.verticalAlign = 'top'  // default value

        }
      }

      Handsontable.renderers.TextRenderer.apply(this, [instance, td, row, col, prop, showValue, cellProperties]);
      // console.log('..........');
    }

    // returns the metadata template
    var getMetadataTemplate = function () {
      var metadata = {}
      metadata.dtype = ''; // 'str', 'int', 'float', 'date'
      metadata.value = '';
      metadata.formula = '';
      metadata.format = {}; // 'bold' -> ('bold', 'normal)

      return metadata;
    }

    // Creates cell Metadata for new cells 
    service.createCellMetadataTemplate = function (hot, row, col) {
      const cellProperties = hot.getCellMeta(row, col);

      if (!cellProperties.hasOwnProperty('metadata')) {
        var metadata = getMetadataTemplate();
        cellProperties.metadata = metadata;
      }
    }

    service.applyFormat = function (hot, button, buttonStatus) {
      // debugger;
      // console.log('ButtonStat in boldFunc', buttonStatus);
      var arr = hot.getSelected()

      // If no cells selected
      if (arr === undefined || arr.length === 0) {
        return;
      }

      var partSelection;
      var startRow, startCol, endRow, endCol;
      var cellTD, cellProperties, format;
      for (var i = 0; i < arr.length; i++) {
        partSelection = arr[i]
        startRow = partSelection[0], startCol = partSelection[1], endRow = partSelection[2], endCol = partSelection[3];
        for (var row = startRow; row <= endRow; row++) {
          for (var col = startCol; col <= endCol; col++) {
            cellTD = hot.getCell(row, col);
            cellProperties = hot.getCellMeta(row, col);
            // Format can be applied even to those cells which do not hold any data
            if (!cellProperties.hasOwnProperty('metadata')) {
              service.createCellMetadataTemplate(hot, row, col);
            }
            format = cellProperties.metadata.format;
            // If button is not present then add it or just alter it
            // Set opposite to button status
            if (['textAlignLeft', 'textAlignCenter', 'textAlignRight'].indexOf(button) > -1) {
              if (buttonStatus !== true)
                format['textAlign'] = button;
              else
                format['textAlign'] = 'textAlignLeft'; // default value
            }
            else if (['verticalAlignTop', 'verticalAlignCenter', 'verticalAlignRight'].indexOf(button) > - 1) {
              if (buttonStatus !== true)
                format['verticalAlign'] = button;
              else
                format['verticalAlign'] = 'verticalAlignCenter'; // default value
            }
            else if (['bold', 'italic', 'underline'].indexOf(button) > -1) {
              if (buttonStatus !== true)
                format[button] = true;
              else
                format[button] = false;
            }
            service.renderCell(hot, row, col)
          }
        }
      }
      // console.log('51st Row, 1st Column');
      // console.log(hot.getCellMeta(50, 1));
    }

    // For manually rendering individual cells
    service.renderCell = function (instance, row, col) {
      const td = instance.getCell(row, col);
      const value = instance.getDataAtCell(row, col);
      const cellProperties = instance.getCellMeta(row, col);
      const prop = cellProperties.prop;

      // It won't render things outside of view (as no dom is created for cells outside of view)
      if (td !== null && td !== undefined) {
        service.customRenderer(instance, td, row, col, prop, value, cellProperties);
      }
    }

    // Managing state of the toolbar button (active or inactive)
    service.isSelectionActive = function (hot, button) {
      var buttonStatus = true; // if all selected cells are already true for this then true else false
      var arr = hot.getSelected();
      if (arr === undefined || arr.length === 0) {
        buttonStatus = false;
        return buttonStatus;
      }

      for (var i = 0; i < arr.length; i++) {
        var selection = arr[i]
        var startRow = selection[0], startCol = selection[1], endRow = selection[2], endCol = selection[3];
        for (var row = startRow; row <= endRow; row++) {
          for (var col = startCol; col <= endCol; col++) {
            var cellProperties = hot.getCellMeta(row, col)
            // If metadata is NOT present OR metadata's format DOESN'T have this particular format property OR 
            // if this format property is set to false, set button status to false
            if (!cellProperties.hasOwnProperty('metadata')) {
              buttonStatus = false;
            }
            else if (['bold', 'italic', 'underline'].indexOf(button) > -1) {
              if (!cellProperties.metadata.format.hasOwnProperty(button) || cellProperties.metadata.format[button] === false)
                buttonStatus = false;
            }
            else if (['textAlignLeft', 'textAlignCenter', 'textAlignRight'].indexOf(button) > -1) {
              if (!cellProperties.metadata.format.hasOwnProperty('textAlign') || cellProperties.metadata.format['textAlign'] !== button)
                buttonStatus = false;
            }
            else if (['verticalAlignTop', 'verticalAlignCenter', 'verticalAlignRight'].indexOf(button) > -1) {
              if (!cellProperties.metadata.format.hasOwnProperty('verticalAlign') || cellProperties.metadata.format['verticalAlign'] !== button)
                buttonStatus = false;
            }
            if (buttonStatus === false)
              return buttonStatus;
          }
        }
      }

      return buttonStatus;
    }


  }

  //---------------------------------------------------------------------------------------------------------------------//
  //---------------------------------------------Helpers Start-----------------------------------------------------------//
  //---------------------------------------------------------------------------------------------------------------------//

  // General javascript functions
  Helpers.$inject = [];
  function Helpers() {
    var service = this;

    // clones an object
    // Doesn't work when the object has date objects
    service.clone = function (obj) {
      return JSON.parse(JSON.stringify(obj))
    }

  }


})() 