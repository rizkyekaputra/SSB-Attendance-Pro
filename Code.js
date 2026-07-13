/*======================================================
  SSB Attendance Pro Enterprise
  Version : 2.0
  Module  : Core
======================================================*/


/*======================================================
  START APPLICATION
======================================================*/
function doGet() {

  return HtmlService
    .createTemplateFromFile("Sidebar")
    .evaluate()
    .setTitle("SSB Attendance Pro Enterprise")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/*======================================================
  INCLUDE HTML
======================================================*/
function include(fileName){

  return HtmlService
    .createHtmlOutputFromFile(fileName)
    .getContent();

}