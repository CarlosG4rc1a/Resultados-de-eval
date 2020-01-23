function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Documento')
      .addItem('Reporte', 'charts')

      .addToUi();
}
function charts(){
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getActiveSheet(); 
  
  var rule = hoja.getRange("B5").getDataValidation();
  var prof = hoja.getRange("B5").getValue();
  var seleccion =ss.getSheetByName(prof);
  //var see = seleccion.getRange("B6").setValue(prof);
 //------------------------------------------------------------------- -----------
   var estilo = {};
    estilo[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    estilo[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    estilo[DocumentApp.Attribute.FOREGROUND_COLOR] = '#4a86e8';
    estilo[DocumentApp.Attribute.FONT_SIZE] = 20;
    estilo[DocumentApp.Attribute.BOLD] = true;
var estilo1 = {};
    estilo1[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    estilo1[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    estilo1[DocumentApp.Attribute.FOREGROUND_COLOR] = '#434343';
    estilo1[DocumentApp.Attribute.FONT_SIZE] = 14;
    estilo1[DocumentApp.Attribute.ITALIC] = true ;
    estilo1[DocumentApp.Attribute.BOLD] = false;
var estilo2 = {};
    estilo2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    estilo2[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    estilo2[DocumentApp.Attribute.FOREGROUND_COLOR] = '#434343';
    estilo2[DocumentApp.Attribute.FONT_SIZE] = 18;
    estilo2[DocumentApp.Attribute.BOLD] = true;


// Plantilla de archivo

    var correo = seleccion.getRange("AA1").getValue();
    //var dir = gjuarez@ifp.mx;
    var doc = DocumentApp.create(prof + ' | Evaluación docente ').addViewer(correo);
        //Header
    var ifp =      doc.addHeader().appendParagraph('Instituto Francisco Possenti');
    var frase =    doc.getHeader().appendParagraph('Per Crucem ad lucem ');
    var sec =      doc.getHeader().appendParagraph('Secundaria');
    var nom =      doc.getHeader().appendParagraph('Evaluación docente');
                   doc.appendHorizontalRule();
        //Body
                   doc.getHeader().appendParagraph('Nombre del profesor:  ' + prof);
                   
    ifp.setAttributes(estilo);
    frase.setAttributes(estilo1);
    sec.setAttributes(estilo2);
    nom.setAttributes(estilo2);

    var docUrl = doc.getUrl();
    var docId = doc.getId();
    var docopen = DocumentApp.openById(docId)
    var body = docopen.getBody();
    

    
  seleccion.getRange("E2001").setValue('Muy bueno: ');
  seleccion.getRange("F2001").setFormula("=COUNTIF(F1:F2000;\"Muy bueno\")");
  seleccion.getRange("E2002").setValue('Bueno: ');
  seleccion.getRange("F2002").setFormula("=COUNTIF(F1:F2000;\"Bueno\")");
  seleccion.getRange("E2003").setValue('Regular: ');
  seleccion.getRange("F2003").setFormula("=COUNTIF(F1:F2000;\"Regular\")");
  seleccion.getRange("E2004").setValue('Necesita mejorar: ');
  seleccion.getRange("F2004").setFormula("=COUNTIF(F1:F2000;\"Necesita mejorar\")");
  var grafico1 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("F2001:F2004"))
  .setPosition(2,3,0,0)
  .setOption('title', 'Esta actualizado')
  grafico1 = grafico1.build();
  
  hoja.insertChart(grafico1)
  var chartBlob = grafico1.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//------------------------------------------------  
  seleccion.getRange("G2001").setFormula("=COUNTIF(G1:G2000;\"Muy bueno\")");
  seleccion.getRange("G2002").setFormula("=COUNTIF(G1:G2000;\"Bueno\")");
  seleccion.getRange("G2003").setFormula("=COUNTIF(G1:G2000;\"Regular\")");
  seleccion.getRange("G2004").setFormula("=COUNTIF(G1:G2000;\"Necesita mejorar\")");
  var grafico2 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("G2001:G2004"))
  .setPosition(2,9,0,0)
  .setOption('title', 'Es buen mediador')
  grafico2 = grafico2.build();
  
  hoja.insertChart(grafico2)
  var chartBlob = grafico2.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//------------------------------------------------ 
  seleccion.getRange("H2001").setFormula("=COUNTIF(H1:H2000;\"Muy bueno\")");
  seleccion.getRange("H2002").setFormula("=COUNTIF(H1:H2000;\"Bueno\")");
  seleccion.getRange("H2003").setFormula("=COUNTIF(H1:H2000;\"Regular\")");
  seleccion.getRange("H2004").setFormula("=COUNTIF(H1:H2000;\"Necesita mejorar\")");
  var grafico3 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("H2001:H2004"))
  .setPosition(20,3,0,0)
  .setOption('title', 'Utiliza diferentes recursos didácticos')
  grafico3 = grafico3.build();
  
  hoja.insertChart(grafico3)
  var chartBlob = grafico3.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//-------------------------------------------------
  seleccion.getRange("I2001").setFormula("=COUNTIF(I1:I2000;\"Muy bueno\")");
  seleccion.getRange("I2002").setFormula("=COUNTIF(I1:I2000;\"Bueno\")");
  seleccion.getRange("I2003").setFormula("=COUNTIF(I1:I2000;\"Regular\")");
  seleccion.getRange("I2004").setFormula("=COUNTIF(I1:I2000;\"Necesita mejorar\")");
  var grafico4 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("I2001:I2004"))
  .setPosition(20,9,0,0)
  .setOption('title', 'Tiene buena comunicación con el grupo')
  grafico4 = grafico4.build();
  
  hoja.insertChart(grafico4)
  var chartBlob = grafico4.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//---------------------------------------------------
  seleccion.getRange("J2001").setFormula("=COUNTIF(J1:J2000;\"Muy bueno\")");
  seleccion.getRange("J2002").setFormula("=COUNTIF(J1:J2000;\"Bueno\")");
  seleccion.getRange("J2003").setFormula("=COUNTIF(J1:J2000;\"Regular\")");
  seleccion.getRange("J2004").setFormula("=COUNTIF(J1:J2000;\"Necesita mejorar\")");
  var grafico5 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("J2001:J2004"))
  .setPosition(39,3,0,0).setOption('title', 'Promueve el trabajo colaborativo')
  grafico5 = grafico5.build();
  
  hoja.insertChart(grafico5)
  var chartBlob = grafico5.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//-------------------------------------------------
  seleccion.getRange("K2001").setFormula("=COUNTIF(K1:K2000;\"Muy bueno\")");
  seleccion.getRange("K2002").setFormula("=COUNTIF(K1:K2000;\"Bueno\")");
  seleccion.getRange("K2003").setFormula("=COUNTIF(K1:K2000;\"Regular\")");
  seleccion.getRange("K2004").setFormula("=COUNTIF(K1:K2000;\"Necesita mejorar\")");
  var grafico6 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("K2001:K2004"))
  .setPosition(39,9,0,0)
  .setOption('title', 'Utiliza lenguaje adecuado')
  grafico6 = grafico6.build();
  
  hoja.insertChart(grafico6)
  var chartBlob = grafico6.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//---------------------------------------------------
  seleccion.getRange("L2001").setFormula("=COUNTIF(L1:L2000;\"Muy bueno\")");
  seleccion.getRange("L2002").setFormula("=COUNTIF(L1:L2000;\"Bueno\")");
  seleccion.getRange("L2003").setFormula("=COUNTIF(L1:L2000;\"Regular\")");
  seleccion.getRange("L2004").setFormula("=COUNTIF(L1:L2000;\"Necesita mejorar\")");
  var grafico7 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("L2001:L2004"))
  .setPosition(58,3,0,0)
  .setOption('title', 'Se asegura de que todos los alumnos aprendan')
  grafico7 = grafico7.build();
  
  hoja.insertChart(grafico7)
  var chartBlob = grafico7.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//---------------------------------------------------  
  seleccion.getRange("M2001").setFormula("=COUNTIF(M1:M2000;\"Muy bueno\")");
  seleccion.getRange("M2002").setFormula("=COUNTIF(M1:M2000;\"Bueno\")");
  seleccion.getRange("M2003").setFormula("=COUNTIF(M1:M2000;\"Regular\")");
  seleccion.getRange("M2004").setFormula("=COUNTIF(M1:M2000;\"Necesita mejorar\")");
  var grafico8 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("M2001:M2004"))
  .setPosition(58,9,0,0)
  .setOption('title', 'Despierta el interes de los alumnos')
  grafico8 = grafico8.build();
  
  hoja.insertChart(grafico8)
  var chartBlob = grafico8.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//----------------------------------------------------------------------
  seleccion.getRange("N2001").setFormula("=COUNTIF(N1:N2000;\"Muy bueno\")");
  seleccion.getRange("N2002").setFormula("=COUNTIF(N1:N2000;\"Bueno\")");
  seleccion.getRange("N2003").setFormula("=COUNTIF(N1:N2000;\"Regular\")");
  seleccion.getRange("N2004").setFormula("=COUNTIF(N1:N2000;\"Necesita mejorar\")");
  var grafico9 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("N2001:N2004"))
  .setPosition(77,3,0,0)
  .setOption('title', 'Es puntal para iniciar y terminar su clase')
  grafico9 = grafico9.build();
  
  hoja.insertChart(grafico9)
  var chartBlob = grafico9.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//-------------------------------------------------------
  seleccion.getRange("O2001").setFormula("=COUNTIF(O1:O2000;\"Muy bueno\")");
  seleccion.getRange("O2002").setFormula("=COUNTIF(O1:O2000;\"Bueno\")");
  seleccion.getRange("O2003").setFormula("=COUNTIF(O1:O2000;\"Regular\")");
  seleccion.getRange("O2004").setFormula("=COUNTIF(O1:O2000;\"Necesita mejorar\")");
  var grafico10 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("O2001:O2004"))
  .setPosition(77,9,0,0)
  .setOption('title', 'Su presencia e imagen es pulcra')
  grafico10 = grafico10.build();
  
  hoja.insertChart(grafico10)
  var chartBlob = grafico10.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//---------------------------------------------------------
  seleccion.getRange("P2001").setFormula("=COUNTIF(P1:P2000;\"Muy bueno\")");
  seleccion.getRange("P2002").setFormula("=COUNTIF(P1:P2000;\"Bueno\")");
  seleccion.getRange("P2003").setFormula("=COUNTIF(P1:P2000;\"Regular\")");
  seleccion.getRange("P2004").setFormula("=COUNTIF(P1:P2000;\"Necesita mejorar\")");
  var grafico11 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("P2001:P2004"))
  .setPosition(96,3,0,0)
  .setOption('title', 'Tiene buenos modales')
  grafico11 = grafico11.build();
  
  hoja.insertChart(grafico11)
  var chartBlob = grafico11.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//--------------------------------------------------------
  seleccion.getRange("Q2001").setFormula("=COUNTIF(Q1:Q2000;\"Muy bueno\")");
  seleccion.getRange("Q2002").setFormula("=COUNTIF(Q1:Q2000;\"Bueno\")");
  seleccion.getRange("Q2003").setFormula("=COUNTIF(Q1:Q2000;\"Regular\")");
  seleccion.getRange("Q2004").setFormula("=COUNTIF(Q1:Q2000;\"Necesita mejorar\")");
  var grafico12 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("Q2001:Q2004"))
  .setPosition(96,9,0,0)
  .setOption('title', 'Devuelve los examenes y realiza retroalimentación')
  grafico12 = grafico12.build();
  
  hoja.insertChart(grafico12)
  var chartBlob = grafico12.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//----------------------------------------------------------
  seleccion.getRange("R2001").setFormula("=COUNTIF(R1:R2000;\"Muy bueno\")");
  seleccion.getRange("R2002").setFormula("=COUNTIF(R1:R2000;\"Bueno\")");
  seleccion.getRange("R2003").setFormula("=COUNTIF(R1:R2000;\"Regular\")");
  seleccion.getRange("R2004").setFormula("=COUNTIF(R1:R2000;\"Necesita mejorar\")");
  var grafico13 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("R2001:R2004"))
  .setPosition(115,3,0,0)
  .setOption('title', 'Informa a los alumnos sobre sus calificaciones')
  grafico13 = grafico13.build();
  
  hoja.insertChart(grafico13)
  var chartBlob = grafico13.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//-----------------------------------------------------------
  seleccion.getRange("S2001").setFormula("=COUNTIF(S1:S2000;\"Muy bueno\")");
  seleccion.getRange("S2002").setFormula("=COUNTIF(S1:S2000;\"Bueno\")");
  seleccion.getRange("S2003").setFormula("=COUNTIF(S1:S2000;\"Regular\")");
  seleccion.getRange("S2004").setFormula("=COUNTIF(S1:S2000;\"Necesita mejorar\")");
  var grafico14 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("S2001:S2004"))
  .setPosition(115,9,0,0)
  .setOption('title', 'Pasa lista')
  grafico14 = grafico14.build();
  
  hoja.insertChart(grafico14)
  var chartBlob = grafico14.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//----------------------------------------
  seleccion.getRange("T2001").setFormula("=COUNTIF(T1:T2000;\"Muy bueno\")");
  seleccion.getRange("T2002").setFormula("=COUNTIF(T1:T2000;\"Bueno\")");
  seleccion.getRange("T2003").setFormula("=COUNTIF(T1:T2000;\"Regular\")");
  seleccion.getRange("T2004").setFormula("=COUNTIF(T1:T2000;\"Necesita mejorar\")");
  var grafico15 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("T2001:T2004"))
  .setPosition(134,3,0,0)
  .setOption('title', 'Reporta indisciplinas')
  grafico15 = grafico15.build();
  
  hoja.insertChart(grafico15)
  var chartBlob = grafico15.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//-----------------------------------------------------
  seleccion.getRange("U2001").setFormula("=COUNTIF(U1:U2000;\"Muy bueno\")");
  seleccion.getRange("U2002").setFormula("=COUNTIF(U1:U2000;\"Bueno\")");
  seleccion.getRange("U2003").setFormula("=COUNTIF(U1:U2000;\"Regular\")");
  seleccion.getRange("U2004").setFormula("=COUNTIF(U1:U2000;\"Necesita mejorar\")");
  var grafico16 = hoja.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(seleccion.getRange("E2001:E2004"))
  .addRange(seleccion.getRange("U2001:U2004"))
  .setPosition(134,9,0,0)
  .setOption('title', 'Mantiene el orden en clase')
  grafico16 = grafico16.build();
  
  hoja.insertChart(grafico16)
  var chartBlob = grafico16.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
//--------------------------------------------------------
//--------------------------------------------------------  
    seleccion.getRange("AA2").setValue('Promedio: ');
  seleccion.getRange("AB2").setFormula("=AVERAGE(W2:W2000)");
  var grafico19 = hoja.newChart()
  .setChartType(Charts.ChartType.GAUGE)
  .addRange(seleccion.getRange('AA2:AB2'))
  .setPosition(176,6,0,0)
  .setOption('height', 300)
  .setOption('width', 300)
  .setOption('title', 'Promedio')
  .setOption('max',10)
  grafico19 = grafico19.build();
  
  hoja.insertChart(grafico19)
  var chartBlob = grafico19.getAs('image/png').copyBlob();
  body.appendImage(chartBlob);
  
//------------------------------------------------------------
 var prop = seleccion.getRange("V1:V").getValues();
 body.appendParagraph('   ' + prop + '   ');
 var nota = seleccion.getRange("X1:X").getValues();
 body.appendParagraph('   ' + nota + '   ');
 
 var content = 'Evaluación docente ' +docUrl;
 
 Utilities.sleep(5000);
 GmailApp.sendEmail(correo,"Evaluación Docente", content);
 }
  

