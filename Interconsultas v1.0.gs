/* 
  
  ANOTAÇÕES
  
  NAVEGUE PELA PLANILHA SE GUIANDO COMO UMA MATRIZ:...
  VAR.FUNC(PARAMENTROS)[LINHA][COLUNA];
  
  ASPAS DUPLAS PARA TEXTO.
  SIMBOLO DE SOMA(+) PARA CONTATENAR.

  
*/



//Função que preenche a coluna Status
function DefineStatus(row){
  var status = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //Ativa a Planilha/aba atual MAS contado de matriz inicia em 1 APENAS nessa função
  var celula = status.getRange(row+1,15); //Define o intervalo STATUS
  celula.setValue("ENVIADO"); //Seta valor ENVIADO no intervalo definido
}

//Função de Alerta
function Alert(cont){
  var interface = SpreadsheetApp.getUi(); //variavel de interface
  
  if (cont > 0){
    interface.alert("SUCESSO! ENVIADO(S) " + cont + " NOVO(S) EMAIL(S)"); //modulo de alerta
  } else {
    interface.alert("NÃO HÁ NOVAS SOLICITAÇÕES DE INTERCONSULTAS CADASTRADAS"); //modulo de alerta
  }
}

//Função que preenche o email
function Email_Body(row, planilha){
  var assunto = "NAC INFORMA SOLICITAÇÃO DE INTERCONSULTA";
  var format_date_solicitacao = Utilities.formatDate(new Date (planilha.getValues()[row][0]), "GMT-03:00", "dd/MM/yyyy HH:mm:ss");
  var format_date_programacao = Utilities.formatDate(new Date (planilha.getValues()[row][10]), "GMT-03:00", "dd/MM/yyyy HH:mm:ss");
  var corpo_estranho;
  
  //Verifica SE a coluna H está vazia, se estiver vazio escreve o valor "Não" na variavel corpo_estranho
  if ( (planilha.getValues()[row][7] == null) || (planilha.getValues()[row][7] == "") ) {
    corpo_estranho = "Não";
  } else {
    corpo_estranho = planilha.getValues()[row][7];
  }
  
  const raw_email = {
    1:"NAC INFORMA SOLICITAÇÃO DE INTERCONSULTA\n",
    2:"\nData da Solicitação: " + format_date_solicitacao,
    3:"\nNome do Paciente: " + planilha.getValues()[row][1],
    4:"\nProntuário: " + planilha.getValues()[row][2],
    5:"\nSetor Solicitante: " + planilha.getValues()[row][4],
    6:"\nDescrição Interconsulta: " + planilha.getValues()[row][6],
    7:"\nCorpo Estranho?: " + corpo_estranho,//planilha.getValues()[row][7],
    8:"\nUrgência: " + planilha.getValues()[row][8],
    9:"\nProgramação: " + format_date_programacao,
  }
  var email = "";
  for(var txt = 1; txt < 10; txt++)
    email = email + raw_email[txt];
  var email_address = planilha.getValues()[row][13];
  
  //Função do GOOGLE API para o envio de email
  MailApp.sendEmail(email_address,assunto,email,{noReply:false});
}


//Função Principal
function Email_Interconsulta(){
  
  //Define a planilha ativa
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var planilha = sheet.getDataRange();
  
  //contador de execuções para função Alert();
  var cont = 0;
  
  //Verifica se o ESPECIALIDADE está preenchido, e utilizar um valor vazio como ponto de parada do Loop
  for (var row = 1; planilha.getValues()[row][5] != ""; row++){
  
  //se STATUS está vazio então chama função Email() e DefineStatus() 
  if ( (planilha.getValues()[row][14] == null) || (planilha.getValues()[row][14] == "") ) {
    Email_Body(row, planilha);
    DefineStatus(row);
    cont++;
   } //FIM SE
  
  } //FIM PARA
  
  Alert(cont);
  
  SpreadsheetApp.flush();//Garante a execução do código ignorando possivel cache

}//Fecha Função Enviar_Email
