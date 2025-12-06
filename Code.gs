/**
 * Gera um Identificador Único Universal (UUID) na versão 4.
 *
 * @return {string} O UUID gerado.
 * @customfunction
 */
function UUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

/**
 * @OnlyCurrentDoc
 *
 * O objeto especial 'e' usado em uma função doGet(e) ou doPost(e) representa um
 * parâmetro de evento que contém informações sobre quaisquer parâmetros de solicitação.
 *
 * @param {object} e O objeto de parâmetro do evento.
 * @return {GoogleAppsScript.HTML.HtmlOutput} O objeto de saída HTML para a página.
 */
function doGet(e) {
  // Retorna o template HTML principal. O template pode acessar as variáveis aqui definidas.
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Treinamento ILEGAL TEOT 2026')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ID da planilha e nomes das abas
const SPREADSHEET_ID = '1nrDTcYC3qhGlJjVlbJ5MUTKwXT5dbQmm6lhGZItAYvI';
const QUESTIONS_SHEET_NAME = 'questoes';
const ANSWERS_SHEET_NAME = 'respostas';

/**
 * Busca todas as questões da planilha que devem ser incluídas na prova.
 * @returns {Array<Object>} Um array de objetos, onde cada objeto representa uma questão.
 */
function getQuestions() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUESTIONS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    // Pega os cabeçalhos para mapear os dados do objeto dinamicamente
    const headers = data.shift();
    
    // Encontra o índice da coluna "Incluir_na_Prova"
    const includeIndex = headers.indexOf('Incluir_na_Prova');
    if (includeIndex === -1) {
      throw new Error("A coluna 'Incluir_na_Prova' não foi encontrada na planilha.");
    }
    
    // Filtra as questões onde Incluir_na_Prova é TRUE e mapeia para objetos
    const questions = data.filter(row => row[includeIndex] === true).map(row => {
      const questionObj = {};
      headers.forEach((header, index) => {
        // Renomeia a chave 'ID' para 'id' para facilitar o uso no frontend
        const key = header === 'ID' ? 'id' : header;
        questionObj[key] = row[index];
      });
      return questionObj;
    });
    
    return questions;
  } catch (error) {
    // Registra o erro para depuração no Apps Script
    console.error('Erro ao buscar questões: ' + error.toString());
    // Retorna um objeto de erro para o frontend
    return { error: 'Falha ao carregar as questões da planilha. Verifique os logs do servidor.' };
  }
}


/**
 * Recebe os resultados do quiz do frontend e os salva na planilha 'respostas'.
 * @param {string} payloadString O payload JSON stringificado contendo os dados do usuário e as respostas.
 * @returns {Object} Um objeto de sucesso ou erro.
 */
function submitResults(payloadString) {
  try {
    const payload = JSON.parse(payloadString);
    const { user, answers, score, provaId } = payload;
    
    if (!user || !answers || !Array.isArray(answers)) {
       throw new Error("Payload inválido. 'user' e 'answers' são obrigatórios.");
    }

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ANSWERS_SHEET_NAME);
    const timestamp = new Date();
    
    // Itera sobre cada resposta no payload e adiciona uma nova linha na planilha
    answers.forEach(answer => {
      const result = answer.isCorrect ? 'Correta' : 'Incorreta';
      
      const newRow = [
        user.email,         // ID_Residente
        user.name,          // Nome_Residente
        provaId,            // ID_Prova
        answer.questionId,  // ID_Questão
        answer.submittedAnswer, // Resposta_Enviada
        result,             // Resultado
        timestamp           // Data_Hora
      ];
      
      sheet.appendRow(newRow);
    });

    return { success: true, message: 'Resultados enviados com sucesso!' };

  } catch (error) {
    console.error('Erro ao enviar resultados: ' + error.toString());
    return { success: false, message: 'Falha ao salvar os resultados. Tente novamente.' };
  }
}

