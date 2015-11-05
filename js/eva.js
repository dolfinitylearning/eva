/**
 * @file
 * Create the Excel VBA Simulator.
 */

"use strict";

var ExcelVbaAnimator;

(function($) {
  /**
   * Also need to have the HTML template for the visualization in 
   * an element with the id "eva-template".
   * @param {type} $parentElement
   * @param {type} programToRun
   * @param {type} numRowsInSs
   * @param {type} numColsInSs
   * @returns {undefined}
   */
  ExcelVbaAnimator = function($parentElement, programToRun, numRowsInSs, numColsInSs) {
    if ( ! $parentElement || ! programToRun || ! programToRun.reset
         || ! programToRun.sourceCode
         || ! numRowsInSs || ! numColsInSs 
         || isNaN(numRowsInSs) || isNaN(numColsInSs) ) {
       alert("Sorry, you need to give the visualization engine everything " 
             + "it needs to be able to run." );
       return;
    }
    this.$parentElement = $parentElement;
    this.programToRun = programToRun;
    this.numRowsInSs = numRowsInSs;
    this.numColsInSs = numColsInSs;
    this.colorProvider = new ColorProvider();
  };
  /**
   * Standard animation duration (ms).
   */
  ExcelVbaAnimator.prototype.standardDuration = 1000;
  /**
   * Global counter of all steps run. Used to skip animations when the
   * user clicks Step when a step is running.
   */
  ExcelVbaAnimator.prototype.stepsRun = 0;
  /**
   * Flag to show whether animations are to be shown.
   */
  ExcelVbaAnimator.prototype.animationOn = true;
  /**
   * Count of number of steps running an animation.
   */
  ExcelVbaAnimator.prototype.stepsRunning = 0;
  /**
   * Flag set if user clicks Step button afore running step is complete.
   */
  ExcelVbaAnimator.prototype.skippingStep = false;
  /**
   * How to indicate emptiness in animations.
   */
  ExcelVbaAnimator.prototype.emptyVarIndicator = "(Empty)";

  //Allocate memory for a notional variable.
  ExcelVbaAnimator.prototype.allocateMemory = function( nVarName, nVarType ) {
    //Assume nVarName dinna exist, and nVarType is OK.
    //Checked by CPU before calling this function.
    //Create the new memory location.
    var message;
    if ( this.nVarExists(nVarName) ) {
      message = this.t("Variable ")+ nVarName
              + this.t(" already declared.");
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    if ( ! this.nDataTypeOk(nVarType) ) {
      message = this.t("Data type ") + nVarType + this.t(" unknown." );
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    var newElementHtml = "<p class='variable " + nVarType + "' "
      +     "data-variable='" + nVarName + "' "
      +     "data-type='" + nVarType + "'>"
      +   "<span class='name-type-container'>"
      +     "<span class='name'>" + nVarName + "</span> " 
      +     "<span class='type'>" + nVarType + "</span> "
      +   "</span>"
      +   "<span class='value'>"
      +     "<input type='text'>"
      +   "</span>"
      + "</p>";
    this.$parentElement.find(".eva-memory").append(newElementHtml);
    var $newElement = this.$parentElement
            .find(".eva-memory p[data-variable=" + nVarName + "]");
    $newElement.show();
  };
  
  /**
   * Set the value of a notional variable. Halt CPU if
   * variable not found.
   * @param {string} nVarName Name of the variable.
   * @param newValue Value of the variable
   */
  ExcelVbaAnimator.prototype.setMemory = function( nVarName, newValue ) {
    var dataType, message, targetVardomElement;
    if ( ! this.nVarExists(nVarName) ) {
      message = this.t("Variable ") + nVarName + this.t(" not declared.");
      this.errorCpuHalt( message );
      return;
//      throw new Error(
//        this.t("Variable ") + nVarName + this.t(" not declared.")
//      );
    }
    //Check data type.
    dataType = this.getNVarType( nVarName );
    if ( this.nDataTypeNumeric(dataType) && isNaN(newValue) ) {
      message = this.t("Data type mismatch");
      this.errorCpuHalt( message );
      return;
      //      throw new Error( message );
    }
    targetVardomElement = this.getNVarDomElement(nVarName);
    targetVardomElement.val(newValue);
  };
  /**
   * Get the value of a notional variable from memory. Halt CPU if
   * variable not found.
   * @param {string} nVarName Name of the variable.
   * @returns Value of the variable.
   */
  ExcelVbaAnimator.prototype.getMemory = function( nVarName ) {
    var message, domElement;
    if ( ! this.nVarExists(nVarName) ) {
      message = this.t("Variable ") + nVarName 
              + this.t(" not declared.");
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    domElement = this.getNVarDomElement(nVarName);
    return domElement.val();      
  };
  /**
   * Get the DOM element for a notional variable with a given name.
   * @param {String} nVarName Name of the notional variable.
   * @returns {Object} DOM element.
   */
  ExcelVbaAnimator.prototype.getNVarDomElement = function(nVarName) {
    return this.$parentElement.find(
      ".eva-memory p[data-variable=" + nVarName + "] " + ".value input"
    );
  };
  /**
   * Get the type of a notional variable with a given name.
   * @param {String} nVarName Name of the notional variable.
   * @returns {Object} DOM element.
   */
  ExcelVbaAnimator.prototype.getNVarType = function(nVarName) {
    return this.$parentElement.find(
      ".eva-memory p[data-variable=" + nVarName + "] " + ".type"
    ).text();
  };
  /**
   * Check whether a notional data type is OK.
   * @param {String} typeToCheck Data type to check, e.g., string.
   * @returns {Boolean} True if known, else false.
   */
  ExcelVbaAnimator.prototype.nDataTypeOk = function( typeToCheck ) {
    switch ( typeToCheck ) {
      case "integer":
      case "single":
      case "string":
        return true;
    }
    return false;
  };
  /**
   * Is the given data type numeric?
   * @param {String} typeToCheck Name of type to check.
   * @returns {Boolean} True if numeric, else false.
   */
  ExcelVbaAnimator.prototype.nDataTypeNumeric = function( typeToCheck ) {
    switch ( typeToCheck ) {
      case "integer":
      case "single":
      case "long":
      case "double":
        return true;
      case "string":
        return false;
      default:
        var message = "nDataTypeNumeric: " + this.t("Unknown data type: ")
          + typeToCheck;
      this.errorCpuHalt( message );
      return;
//        throw new Error(message);
    }
  };
  //Check whether a notional variable exists.
  ExcelVbaAnimator.prototype.nVarExists = function(nVarName) {
    var domElement = this.getNVarDomElement(nVarName);
    return ( domElement.length == 1 );
  };

    
  ExcelVbaAnimator.prototype.showWhatHappened = function( message ) {
     var text, logDisplay;
     text = this.t( message );
     logDisplay = this.$parentElement.find(".eva-event-log");
     logDisplay.html(text).scrollTop(0);
     return logDisplay.css({ opacity: 0 })
       .animate( { opacity: 1 } )
       .delay(750)
       .promise();
   };
   /** Flag to show whether the CPU has halted, either due to a runtime error,
    *  or normal program termination.
    */
  ExcelVbaAnimator.prototype.cpuHalted = false;
  /**
   * Show the user a message from the CPU.
   * @param {String} message Message to show.
   */
  ExcelVbaAnimator.prototype.showCpuMessage = function( message ) {
    this.$parentElement.find(".eva-cpu-message").html(message);
  };
  /**
   * Get the number of the next statement to run from the instruction pointer.
   * @returns {Integer} Statement number, or throw error.
   */
  ExcelVbaAnimator.prototype.getNextStatementNumber = function() {
    var statementNumber = this.$parentElement
      .find(".eva-cpu-components .eva-next-instruction input").val();
    this.checkStatementNumber( statementNumber );
    return statementNumber;
  };
  /**
   * Set the next instruction number.
   * @param {Integer} statementNumber Instruction number.
   */
  ExcelVbaAnimator.prototype.setNextStatementNumber = function( statementNumber ) {
    this.$parentElement.find(".eva-cpu-components .eva-next-instruction input").val(statementNumber);
    this.checkStatementNumber( statementNumber );
    return true;
  };
  /**
   * Check statement number.
   * @param {Integer} statementNumber Statement number.
   */
  ExcelVbaAnimator.prototype.checkStatementNumber = function(statementNumber) {
    if (     ! statementNumber
          || isNaN(statementNumber)
          || statementNumber < 1
          || ! this.program[statementNumber] ) {
      var message = this.t("Bad statement number: ") + statementNumber;
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
  };
  /**
   * Set the value of the evaluator.
   * @param {String|Number} newValue Value to show.
   */
  ExcelVbaAnimator.prototype.setEvaluator = function( newValue ) {
    var $evaluator = this.$parentElement.find(".eva-cpu-components #evaluator");
    $evaluator.text( newValue );
  };
  /**
   * Declare a notional variable. Sets cpu.halt if there's a problem.
   * @param {string} nVarName Name, e.g., i
   * @param {string} nVarType Type, e.g., integer
   * @param {string} note Note to show the user
   * @returns True if no error, else false.
   */
  ExcelVbaAnimator.prototype.declareNVar = function( nVarName, nVarType, note) {
    var message;
    if ( this.memory.nVarExists(nVarName) ) {
      message = this.t("Variable ") + nVarName 
              + this.t(" already declared.");
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    if ( ! this.memory.nDataTypeOk(nVarType) ) {
      message = this.t("Data type ") + nVarType 
              + this.t(" unknown.");
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    this.allocateMemory(nVarName, nVarType);
    return true;
  };
  /**
   * Get the DOM element used for the evaluator.
   * @returns {undefined}
   */
  ExcelVbaAnimator.prototype.getEvaluatorDomElement = function() {
    return this.$parentElement.find(".eva-cpu-components #evaluator");
  };
  /**
   * Notional machine error. Halt the notional program.
   * @param {string} message Error message.
   */
  ExcelVbaAnimator.prototype.errorCpuHalt = function(message) {
    //Translate and show the message.
    message = "<span class='runtime-error'>" + this.t(message);
    message += " System halted. Click Reset to start over." + "</span>";
    this.showCpuMessage(message);
    //Set halted flag.
    this.cpuHalted = true;
    //Set state of run controls.
    this.setControlState("halted");
  };
  /**
   * Program halts normally.
   */
  ExcelVbaAnimator.prototype.normalHalt = function() {
    this.cpuHalted = true;
    this.showCpuMessage("Normal halt");
    this.$parentElement
      .find(".eva-cpu-components .eva-next-instruction").val("");
    this.highlightStatement("none");
    this.setControlState("halted");
  };
  ExcelVbaAnimator.prototype.runNextStatement = function(  ) {
    var statementNumber, currentStatement, thisy;
    $(".movey-thing").stop().hide();
    this.stepsRun ++;
    if ( this.stepsRunning > 0 ) {
      window.setTimeout(50, this.runNextStatement);
      return;
    }
    //Fetch the next statement to run.
    statementNumber = this.getNextStatementNumber();
    //Fetch statement.
    currentStatement = this.program[statementNumber];
    currentStatement.stepNumber = this.stepsRun;
    thisy = this;
    try {
      currentStatement.animation();
    }
    catch(ex){
      thisy.errorCpuHalt(ex.message);
    }
  }; //End runNextStatement.
  
  ExcelVbaAnimator.prototype.move2NextStatement = function(destinationLabel) {
    var statementNumber, message;
    if ( typeof destinationLabel == "undefined" ) {
      //No label give. Advance to next statement.
      statementNumber = this.getNextStatementNumber();
      //Could be a bad statement number.
      if ( this.cpuHalted ) {
        return;
      }
      statementNumber ++;
      if ( statementNumber > this.getNumberStatements() ) {
        //Reached the end of the program.
        this.normalHalt();
        return;
      }
    }
    else {
      //A label was passed.
      statementNumber 
        = this.getStatementNumberWithLabel( destinationLabel );
      if ( destinationLabel == -1 ) {
        message = this.t("Label '") + destinationLabel + this.t("' not found");
        this.errorCpuHalt( message );
        return;
//        throw new Error( message );
      }
    }
    this.highlightStatement( statementNumber );
    this.setNextStatementNumber( statementNumber );
  };
  
  //Array with the program. Each element is a statement.
  ExcelVbaAnimator.prototype.program = null; 

  /**
   * Append a program statement.
   * @param {Integer} statementNumber Statement number.
   * @param {String} statementSourceCode Source code. If includes {{, then
   *        a span tag is added for an expression to highlight.
   */
  ExcelVbaAnimator.prototype.appendStatementToDisplay = function(statementNumber, statementSourceCode) {
    var html;
    statementSourceCode 
            = statementSourceCode.replace("{{", "<span class='expression'>");
    statementSourceCode 
            = statementSourceCode.replace("}}", "</span>");
    html = "<li data-statement-number='" + statementNumber + "'><pre>" 
            + statementSourceCode + "</pre></li>";
    this.$parentElement.find(".eva-code ol").append(html);
  };
  /**
   * Highlight a statement, or no statement.
   * @param {Integer|String} statementNumber Statement number, or "none"
   *        to highlight nothing.
   */
  ExcelVbaAnimator.prototype.highlightStatement = function( statementNumber ) {
    this.$parentElement
            .find(".eva-code ol li pre")
            .removeClass("current-statement");
    if ( statementNumber != "none" ) {
      this.$parentElement.find(
        ".eva-code ol li[data-statement-number=" + statementNumber  + "] pre"
      )
        .addClass("current-statement");
    }
  };
  /**
   * Load a program. 
   * @param {Array of objects} programToLoad Statements.
   */
  ExcelVbaAnimator.prototype.loadProgam = function( programToLoad ) {
    var statementNumber = 0;
    var statement;
    this.program = new Array();
    for( var index in programToLoad ) {
      statementNumber ++;
      statement = programToLoad[index];
      //Give the statement a number.
      statement.statementNumber = statementNumber;
      //Remember the statement.
      this.program[statementNumber] = statement;
      //Add it to the program display.
      this.appendStatementToDisplay(statementNumber, statement.statementSourceCode);
    }
  };
  /**
   * Get the number of statements in the program.
   * @returns {integer} Number of statements.
   */
  ExcelVbaAnimator.prototype.getNumberStatements = function() {
    return this.program.length - 1;
  };
  /**
   * Get the DOM element for the expression in a statement with a 
   * given label.
   * @param {type} statementLabelToFind Label to look for.
   * @returns {DOM element} DOM element.
   */
  ExcelVbaAnimator.prototype.getDomElementExpressionInStatement = function( statementLabelToFind ) {
    var index, statmnt, selector, elemnt, message;
    for ( index in this.program ) {
      statmnt = this.program[index];
      if ( statmnt.label && statmnt.label == statementLabelToFind ) {
        selector = ".eva-code li[data-statement-number=" 
                + statmnt.statementNumber + "] .expression";
        elemnt = this.$parentElement.find(selector);
        if ( elemnt.length == 0 ) {
          message = this.t(
            "getDomElementExpressionInStatement: "
            + "Cannot find element: ") + selector;
          this.errorCpuHalt( message );
          return;
//          throw new Error( message );
        }
        return elemnt;
      }
    }
    message = this.t(
            "getDomElementExpressionInStatement: "
            + "Cannot find label: " + statementLabelToFind);
      this.errorCpuHalt( message );
      return;
//    throw new Error( message );
  };
  /**
   * Find the number of the statement with a given label.
   * @param {String} labelToFind Label to find.
   * @returns {Number} Statement number, or -1 if not found.
   */
  ExcelVbaAnimator.prototype.getStatementNumberWithLabel = function( labelToFind ) {
    var index, statementToCheck;
    for (var index in this.program) {
      statementToCheck = this.program[index];
      if ( statementToCheck.label && statementToCheck.label == labelToFind ) {
        //Found it.
        return statementToCheck.statementNumber;
      }
    }
    //Didn't find it.
    return -1;
  };
  
  /**
   * Stuff about worksheets.
   */
  
  /**
   * Create spreadsheet HTML.
   * @param {Integer} rows Number of rows.
   * @param {Integer} cols Number of columns.
   */
  ExcelVbaAnimator.prototype.createSpreadsheetHtml = function( rows, cols ) {
    var htmlCode, i, j;
    htmlCode = "<table>\n <tr>\n  <th></th>\n";
    for ( i = 65; i < (65 + cols - 1); i++ ) {
      htmlCode += "  <th>" + String.fromCharCode(i) + "</th>\n";
    }
    htmlCode += " </tr>\n";
    for ( i = 1; i <= rows; i++ ) {
      htmlCode += " <tr data-row='" + i + "'>\n";
      htmlCode += "  <th>" + i + "</th>\n";
      for ( j = 1; j <= cols; j++ ) {
        htmlCode += "  <td data-col='" + j + "'><input type='text'></td>\n";
      }
      htmlCode += " </tr>\n";
    }
    htmlCode += "</table>\n";
    return htmlCode;
  };
    
  /**
   * Get the DOM element for the table that has the spreadsheet.
   * @returns {$}
   */
  ExcelVbaAnimator.prototype.getSsElement = function() {
    return $( this.$parentElement.find(".eva-spreadsheet table") );
  };
  
  /**
   * Set a cell's value. 
   * @param {Integer} row Row.
   * @param {Integer} col Column.
   * @param newValue Value to set into cell.
   * @returns {Boolean} True if no problems, else false.
   */
  ExcelVbaAnimator.prototype.setCellValue = function( row, col, newValue ) {
    var $domCell, message;
    $domCell = this.getCellDomElement(row, col);
    if ( ! $domCell ) {
      message = this.t( 
        "setCellValue: bad cell address: " + row + " " + col
      );
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    $domCell.val(newValue);
    return true;
  };
  /**
   * Get the value in a cell.
   * @param {Integer} row Row.
   * @param {Integer} col Column.
   * @returns {Boolean} True if OK, false if bad row,col.
   */
  ExcelVbaAnimator.prototype.getCellValue = function( row, col ) {
    var $domCell, cellValue, message;
    $domCell = this.getCellDomElement(row, col);
    if ( ! $domCell ) {
      message = this.t( 
        "getCellValue: bad cell address: " + row + " " + col
      );
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    cellValue = $domCell.val();
    if ( cellValue == "" ) {
      cellValue = this.emptyVarIndicator;
    }
    return cellValue;
  };
  /**
   * Get the DOM element for a cell.
   * @param {Integer} row Row.
   * @param {Integer} col Column.
   * @returns {Boolean|DOM element} Elements, or false if not 
   *  valid cell coords.
   */
  ExcelVbaAnimator.prototype.getCellDomElement = function( row, col) {
    var $domRow, $domCell;
    row = $.trim(row);
    $domRow = this.$parentElement.find(
            ".eva-spreadsheet tr[data-row=" + row + "]");
    if ( $domRow.length == 0 ) {
      this.errorCpuHalt( "Spreadsheet row " + row + " unknown." );
      return false;
    }
    col = $.trim(col);
    $domCell = $domRow.find("td[data-col=" + col + "] input");
    if ( $domCell.length == 0 ) {
      this.errorCpuHalt( "Spreadsheet column " + col + " unknown." );
      return false;
    }
    return $domCell;
  };
  
  /**
   * Set the state of the controls for user to run the program.
   * @params {String} stateName Name of the state. 
   */
  ExcelVbaAnimator.prototype.setControlState = function(stateName){
    switch(stateName) {
      case "halted":
        this.$parentElement.find(".eva-controls .eva-step").prop("disabled",true);
        this.$parentElement
          .find(".eva-cpu-components .eva-run-status").text("Halted");
        //Turn on Reset button.
        this.$parentElement.find(".eva-reset").prop('disabled', false);
        break;
      case "running":
        this.$parentElement.find(".eva-controls .eva-step")
          .prop("disabled",false);
        this.$parentElement.find(".eva-cpu-components .eva-run-status")
          .text("Running");
        break;
      default:
        this.errorCpuHalt("Unknown runControls state: " + stateName);
    }
  };
  
  /**
   * Reset the visualization to its initial state.
   */
  ExcelVbaAnimator.prototype.reset = function( ) {
    $(".movey-thing").stop().hide();
    this.cpuReset();
    //Clear memory.
    this.$parentElement.find(".eva-memory").html("");
    //Erase everything from the spreadsheet.
    this.$parentElement.find(".eva-spreadsheet td input").val("");
    //Call program reset.
    this.programToRun.reset( this );
  };
  /**
   * Reset the CPU, to get ready to start the program.
   */
  ExcelVbaAnimator.prototype.cpuReset = function() {
    //Clear message.
    this.$parentElement.find(".eva-cpu-message").html("");
    //Clear explanation.
    this.$parentElement.find(".eva-cpu-components .eva-event-log").html("");
    //Clear evaluator.
    this.$parentElement.find(".eva-cpu-components #evaluator").text(
      this.emptyVarIndicator 
    );
    //Set next instruction.
    this.$parentElement.find(".eva-cpu-components .eva-next-instruction input").val(1);
    this.highlightStatement(1);
    //Set running controls state.
    this.setControlState("running");
    //Show running.
    this.cpuHalted = false;
  };
  /**
   * Animate the movement of data. Could be more than one source, e.g.,
   * three notional variables move to evaluator at the same time.
   * There is only one destination.
   * @param {Object|Array of objects} sourceObjects Where data comes from.
   * @param {Object} destObject Where the data goes to.
   * @return {Object} Promise for last animation.
   */
  ExcelVbaAnimator.prototype.animateDataMovement = function(sourceObjects, destObject) {
    //Destination could be a notional variable, a cell, 
    //or the evaluator.
    var source, index, promise, message;
    if ( ! sourceObjects ) {
      message = "animateDataMovement: missing source objects.";
      this.errorCpuHalt( message );
      return;
//      throw new Error("animateDataMovement: missing source objects.");
    }
    if ( ! destObject ) {
      message = "animateDataMovement: missing destination object.";
      this.errorCpuHalt( message );
      return;
//      throw new Error("animateDataMovement: missing destination object.");
    }
    //Iterate through the source objects, creating animations to destination
    for (index in sourceObjects) {
      source = sourceObjects[index];
      promise = this.moveFieldValue( source, destObject );
    }
    //Return the last promise made.
    return promise;
  };
  /**
   * Compute the DOM element used for a source or destination of an
   * animation, given a spec for it. 
   * @param {Object} spec Spec for source/destination object.
   * @returns {Object} DOM element.
   */
  ExcelVbaAnimator.prototype.computeDomElementFromSpec = function( spec ) {
    var domElement, message;
    switch (spec.elementType) {
      case "nVar":
        if ( ! spec.name ) {
          message = this.t("computeDomElementFromSpec: missing nVar name.");
          this.errorCpuHalt( message );
          return;
//          throw new Error(
//            this.t("computeDomElementFromSpec: missing nVar name.")
//          );
        }
        domElement = this.getNVarDomElement(spec.name);
        break;
      case "cell":
        if ( ! spec.row ) {
          message = this.t("computeDomElementFromSpec: missing row.");
          this.errorCpuHalt( message );
          return;
//          throw new Error(
//            this.t("computeDomElementFromSpec: missing row.")
//          );
        }
        if ( ! spec.col ) {
          message = this.t("computeDomElementFromSpec: missing col.");
          this.errorCpuHalt( message );
          return;
//          throw new Error(
//            this.t("computeDomElementFromSpec: missing col.")
//          );
        }
        domElement = this.getCellDomElement(
                  spec.row,
                  spec.col
               );
        break;
      case "evaluator":
        domElement = this.getEvaluatorDomElement();
        break;
      case "expressionInStatement":
        //Choose the marked expression in a code statement.
        if ( ! spec.statementLabel ) {
          message = this.t("computeDomElementFromSpec: missing label.");
          this.errorCpuHalt( message );
          return;
//          throw new Error(
//            this.t("computeDomElementFromSpec: missing label.")
//          );
        }
        domElement = this.getDomElementExpressionInStatement( 
          spec.statementLabel
        );
        break;
      default:
        message = this.t("computeDomElementFromSpec: unknown element type: ");        
        this.errorCpuHalt( message );
        return;
//        throw new Error(
//            this.t("computeDomElementFromSpec: unknown element type: ") 
//            + spec.elementType
//        );
    }
    return domElement;
  };
  /**
   * Animated move of the value of a source to a destination.
   * @param {Object} sourceElement Coords of source: {top:??, left:??}
   * @param {Object} destElement Coords of destination: {top:??, left:??}
   * @param {Object} statement Statement this is for.
   * @param {String|Number} valueToShow Value to show moving. Omit to show
   *        the value of the source.
   */
  ExcelVbaAnimator.prototype.moveFieldValue = function( sourceElement, destElement, 
          statement, valueToShow ) {
    var message, sourceOffset, destOffset, midpointOffset, sourceTagType;
    var colorToUse, fauve, thisythis;
    if ( ! statement || ! statement.stepNumber ) {
      message = this.t("moveFieldValue: Statement not set correctly.");
      this.errorCpuHalt( message );
      return;
//      throw new Error(message);
    }
    //Compute midpoint of line from source to dest.
    sourceOffset = sourceElement.offset();
    destOffset = destElement.offset();
    midpointOffset = this.computeMidpoint(sourceOffset, destOffset);
    //Set the value to show.
    if (typeof valueToShow == 'undefined') {
      sourceTagType = sourceElement.prop("tagName").toLowerCase();
      valueToShow = (sourceTagType == "input")
                  ? sourceElement.val() 
                  : sourceElement.text();
    }
    //Make a faux to move - a fauve.
    colorToUse = this.colorProvider.getColor();
    fauve = $(
        "<div class='movey-thing' style='"
            + "color: #" + colorToUse + ";"
            + "opacity:1;"
            + "font-size:100%;"
            + "position: absolute;"
            + "vertical-align: top;"
            + "left: " + sourceOffset.left + "px;"
            + "top: " + sourceOffset.top + "px;"
            + "text-shadow: 3px 3px 2px rgba(200, 200, 200, 0.5);"
            +  "'>" + valueToShow + "</div>"
    );
    $("body").append(fauve);
    //Position the fauve.
    fauve.offset( sourceOffset );
    //Start moving animation.
    fauve.show();
    //Make a ref this to this to be used in lambdas.
    thisythis = this;
    return fauve.animate(
      {
        fontSize: "200%",
        opacity: 0.5,
        left: midpointOffset.left,
        top: midpointOffset.top,
        easing: "linear"
      },
      {
        queue: false,
        duration: thisythis.computeDuration( statement ),
        complete: function() {
          fauve.css("position", "absolute")
            .offset(midpointOffset);
          fauve.animate( 
            {
              fontSize: "100%",
              opacity: 0.1,
              top: destOffset.top,
              left: destOffset.left,
              easing: "linear"
            },
            {
              queue: false,
              duration: thisythis.computeDuration( statement ),
              complete: function() {
                //Hide the movey thing.
                fauve.hide();
              }
            }
          );
        } //End second animation.
      }
    ).promise(); //End animate.
  };
  
  
  ExcelVbaAnimator.prototype.initStep = function(){
    //Remove left over DOM elements created during animations.
    this.cleanUpMoveyThings();
    //Hack: disable reset button.
    this.$parentElement.find(".eva-reset").prop('disabled', 'disabled');
    this.$parentElement.find(".eva-cpu-container .eva-event-log")
      .html("");
    this.stepsRunning ++;
  };
  
  
  ExcelVbaAnimator.prototype.finishStep = function() {
    //Hack: enable reset button.
    this.$parentElement.find(".eva-reset").prop('disabled', false);
    this.stepsRunning --;
  };
  
  /**
   * Hackey function to remove DOM elements created during animations.
   * They need to stay existing, until all their promises are finished.
   */
  ExcelVbaAnimator.prototype.cleanUpMoveyThings = function() {
    $(".movey-thing").stop().clearQueue().remove();
  };
  /**
   * Compute the midpoint of a line.
   * @param {Object} source Left,top
   * @param {Object} dest Left,top
   * @returns {Object} Left,top
   */
  ExcelVbaAnimator.prototype.computeMidpoint = function( source, dest ) {
    //Imagine a large triangle, with height and width the
    //diff betwixt source and dest.
    //Compute hypotenuse of large, divide by 2 to get hyp of 
    //small triangle that starts at source.
    //Use trig to compute width and height of small triangle.
    //Use them as offsets from source, to get midpoint coords.
    var bigTriWidth = source.left - dest.left;
    var bigTriHeight = source.top - dest.top;
    var bigTriHyp = Math.sqrt(
            bigTriWidth * bigTriWidth
          + bigTriHeight * bigTriHeight
        );
    var angle = Math.atan( Math.abs( bigTriHeight / bigTriWidth) );
    var smallTriHyp = bigTriHyp/2;
    var smallTriWidth = Math.cos(angle) * smallTriHyp;
    var smallTriHeight = Math.sin(angle) * smallTriHyp;
    var midPoint = {};
    if ( source.left < dest.left ) {
      midPoint.left = source.left + smallTriWidth;
    }
    else {
      midPoint.left = source.left - smallTriWidth;
    }
    //Top and Y are opposite (Top: 0 is top of axis, Y: o is bottom of avis)
    if ( source.top < dest.top ) {
      midPoint.top = source.top + smallTriHeight;
    }
    else {
      midPoint.top = source.top - smallTriHeight;
    }        
    return midPoint;
  };
  /**
   * Animate one data element.
   * @param {DOM object} elementToFlash Element to flash.
   * @param {Object} statement Statement this is for.
   */
  ExcelVbaAnimator.prototype.flashDomElement = function( elementToFlash, statement ) {
    var message, colorToUse, $overlay, offsetTemp;
    if ( ! statement || ! statement.stepNumber ) {
      message = this.t("flashDomElement: Statement not set correctly.");
      this.errorCpuHalt( message );
      return;
//      throw new Error(message);
    }
    //Overlay a rectangle.
    colorToUse = this.colorProvider.getColor();
    $overlay = $("<p class='movey-thing' style='"
              + "opacity:1;"
              + "font-size:100%;"
              + "border: 4px #" + colorToUse + " solid;"
              + "background-color: #" + colorToUse + ";"
              + "width:" + ( parseInt(elementToFlash.width() + 4) ) + "px;"
              + "height:" + ( parseInt(elementToFlash.height() + 4) ) + "px;"
              + "position: absolute;'>&nbsp;</p>");
    this.$parentElement.find(".eva-memory").append( $overlay );
    offsetTemp = {
      left: elementToFlash.offset().left - 2,
      top: elementToFlash.offset().top - 2
    };
    $overlay.offset( offsetTemp );
    //Animate the rectangle away.
    return $overlay.animate(
        {
          opacity: 0
        },
        {
          duration: this.computeDuration( statement ),
          complete: function() {
            $overlay.hide();
          }
        }
    ).promise();
  };
  
  /**
   * Compute the duration of an animation. If a step is being skipped,
   * return 0.
   * @param {Object} statement Statement the duration is for.
   * @returns {Number} Duration in ms.
   */
  ExcelVbaAnimator.prototype.computeDuration = function( statement ) {
    var message;
    if ( ! statement || ! statement.stepNumber ) {
      message = this.t("computeDuration: Statement not set correctly.");
      this.errorCpuHalt( message );
      return;
//      throw new Error(message);
    }
    return ( statement.stepNumber == this.stepsRun )
      ? this.standardDuration
      : 0;
  };
  
  /**
   * Set up this.t(), the translate function.
   * If Drupal.t() exists, use that. Otherwise don't translate.
   * @param {String} message Message to translate.
   */
  ExcelVbaAnimator.prototype.t = typeof Drupal != "undefined" 
    ? Drupal.t
    : function( message ) { return message; };
  

  ExcelVbaAnimator.prototype.setup = function() {
    var htmlClone, ssHtml, thisythis;
    //Create the HTML from template.
    htmlClone = $("#eva-template").clone();
    //Put HTML into location given.
    this.$parentElement.html( htmlClone.html() );
    ssHtml = this.createSpreadsheetHtml( this.numRowsInSs, this.numColsInSs );
    this.$parentElement.find(".eva-spreadsheet").append( ssHtml );
    //Make a ref this to this to be used in lambdas.
    thisythis = this;
    //Load the program
    this.loadProgam(this.programToRun.sourceCode);
    //Reset the notional machine.
    this.reset();
    this.$parentElement.find(".eva-reset").click( function() {
      thisythis.reset();
    });
    this.$parentElement.find(".eva-step").click( function() {
      thisythis.runNextStatement(); 
    });
    this.$parentElement.find(".eva-memory-container")
      .on("blur", "input", function(){
        thisythis.validateMemoryInput(this);
    });
    this.$parentElement.find(".eva-memory-container")
      .on("keypress", "input", function(event){
        if ( event.which == 13 ) {
          event.preventDefault();
          thisythis.validateMemoryInput(this);
        }
    });
    this.$parentElement.find(".eva-next-instruction input").blur(function(){
      thisythis.validateNextInstructionInput(this);
    });
    this.$parentElement.find(".eva-next-instruction input")
            .keypress(function(event){
      if ( event.which == 13 ) {
        event.preventDefault();
        thisythis.validateNextInstructionInput(this);
      }
    });
  };
  
  /**
   * Check what the user typed into memory.
   * @param {Object} inputField The input field typed into.
   * @returns {Boolean} True if OK, else false.
   */
  ExcelVbaAnimator.prototype.validateMemoryInput = function( inputField ) {
    var container, message, $inputField, inputValue;
    //Test the variable's data type.
    $inputField = $(inputField);
    inputValue = $(inputField).val();
    inputValue = $.trim(inputValue);
    $(inputField).val( inputValue );
    message = "";
    container = $(inputField.closest("p"));
    if ( container.hasClass("string") ) {
      //Replace MT string with MT indicator.
      if ( inputValue == "" ) {
        $inputField.val( this.emptyVarIndicator );
      }
    }
    else if ( container.hasClass("integer") || container.hasClass("long") ) {
      //Replace MT string with 0.
      if ( inputValue == "" ) {
        $inputField.val( 0 );
      }
      else if ( isNaN( inputValue ) ) {
        message = this.t("Sorry, you can only enter numbers.");
      }
      else if ( parseInt( inputValue ) != inputValue ) {
        message = this.t("Sorry, you can only enter whole numbers.");
      }
    }
    else if ( container.hasClass("single") || container.hasClass("double") ) {
        //Replace MT string with 0.
        if ( inputValue == "" ) {
          $inputField.val( 0 );
        }
        else if ( isNaN( inputValue ) ) {
          message = this.t(
            "Sorry, you can only enter numbers."
          );
        }
    }
    else {
      message = "validateMemoryInput: " + this.t("Bad data type");
      this.errorCpuHalt( message );
      return;
//      throw new Error( message );
    }
    if ( message == "" ) {
      return true;
    }
    else {
      alert( message );
      $inputField.focus();
      inputField.select();
      return false;
    }
  };
  
  /**
   * Check what the user typed into the instruction pointer.
   * @param {type} inputField The input field typed into.
   * @returns {Boolean} True if OK, else false.
   */
  ExcelVbaAnimator.prototype.validateNextInstructionInput = function( inputField ) {
    var statementNumber, message;
    statementNumber = $(inputField).val();
    statementNumber = $.trim(statementNumber);
    $(inputField).val( statementNumber );
    if (     ! statementNumber
          || isNaN(statementNumber)
          || statementNumber < 1
          || ! this.program[statementNumber] ) {
      message = this.t("Please enter a number between 1 and ")
        + this.getNumberStatements() + ".";
      alert(message);
      $(inputField).focus();
      inputField.select();
      return false;
    }
    this.setNextStatementNumber( statementNumber );
    this.highlightStatement( statementNumber );
    return true;    
  };

  /**
   * Create messages for "what just happened."
   */
  ExcelVbaAnimator.prototype.whatHappened = {
    assignToMemory: "The CPU computed the value on the right of the =, "
            + "then put it in the "
            + "variable on the left of the =.",
    assignToCell: "The CPU computed the value on the right of the =, "
            + "then put it in the worksheet "
            + "cell on the left of the =.",
    comment: "The line starts with ', a single quote. That tells the "
          + "CPU to ignore the rest of the line. The line is a comment. "
          + "You normally use comments to explain how the program works, "
          + "so other programmers can change it if they need to.",
    declareVariable: function(varName, dataType) {
      var infoType;
      switch( dataType ) {
        case "string": 
          infoType = "a string (characters you can type)";
          break;
        case "integer":
          infoType = "an integer (a whole number)";
          break;
        case "long":
          infoType = "a long integer (a whole number)";
          break;
        case "single":
          infoType = "a single (a number that can have a fraction, like 3.14)";
          break;
        case "double":
          infoType = "a double (a number that can have a fraction, like 3.14)";
          break;
        default:
          alert("declareVariable: Bad type: " + dataType);
      }
      return "The CPU allocated some memory, and named it \"" + varName + "\". "
            + "It can store " + infoType + ".";
    }
  };

  var ColorProvider = function() {
    
  };
  /**
   * List of colors to cycle through.
   */
  ColorProvider.prototype.colorList = ["872441", "3CBEE0", "8C8F8D", "F18437", "60CC57"];
  //Uncomment for an alternate list.
  //  colorProvider.colorList = ["00FF64", "FFFF00", "FF0058", "FF6F00", "006F94"];
  /**
   * Current color.
   */
  ColorProvider.prototype.currentColorIndex = 0;
  /**
   * Get a color.
   * @returns {String} Hex code for a color.
   */
  ColorProvider.prototype.getColor = function() {
    var colorToUse = this.colorList[this.currentColorIndex];
    this.currentColorIndex++;
    if ( this.currentColorIndex == this.colorList.length ) {
      this.currentColorIndex = 0;
    }
    return colorToUse;
  };
  
}(jQuery));
