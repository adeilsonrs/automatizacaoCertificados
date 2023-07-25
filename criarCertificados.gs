// Salva certificados em formato doc e pdf em uma única pasta

function criarCertificados() {
  // ID da planilha // No exemplo "https://docs.google.com/spreadsheets/d/11YjwJ8t2jAnvwrgrYxQNpdXmrBMRZ4QmF1WHOWSy25w/edit#gid=0", o ID é:
  var planilhaId = '11YjwJ8t2jAnvwrgrYxQNpdXmrBMRZ4QmF1WHOWSy25w';

  try {
    // Abrindo a planilha usando a variavel que acabamos de criar
    var planilha = SpreadsheetApp.openById(planilhaId);
    var aba = planilha.getSheetByName('Pag1'); // Aqui a gente salva o nome da aba em uma variável

    
    var dados = aba.getDataRange().getValues(); // Obtendo dados da planilha

    // Índices das colunas // Se necessitar de mais informações da tabela é só adicionar novas variáveis aqui, no loop e na parte dos marcadores
    var indiceNome = 0; // Coloque o índice da coluna com o nome
    var indiceMatricula = 1; //Coloque o índice da coluna com a matrícula

    // Cria a pasta "Certificados", se ela já não existir
    var pastaPrincipal = DriveApp.getRootFolder();
    var pastaCerts;
    var pastas = pastaPrincipal.getFoldersByName('Certificados'); // Detalhe: Não é Case Sensitive, se houver uma pasta com C minúsculo, vai salvar nela
    if (pastas.hasNext()) {
      pastaCerts = pastas.next();
    } else {
      pastaCerts = pastaPrincipal.createFolder('Certificados');
    }

    // ID do documento do Google Docs //Atente-se para que seja um arquivo doc, nada de apresentação ou outra planilha
    var modeloDocumentoId = '11UC9YT6t_atB1gjZaovpxexrfuyIuoU5bgdgDYXiTQA';

    // Loop através da planilha // Aqui adicionei começando do i = 1 para saltar a primeira linha //Desative os filtros da planilha caso esteja usando ou use se for a intenção
    for (var i = 1; i < dados.length; i++) { 
      var nomeAluno = dados[i][indiceNome];
      var matriculaAluno = dados[i][indiceMatricula];

      // Copia o modelo do documento para criar um novo documento na pasta "Certificados"
      var novoDocumento = DriveApp.getFileById(modeloDocumentoId).makeCopy(pastaCerts);

      // Renomeia o documento com o nome
      novoDocumento.setName(nomeAluno + '_certificado');

      // Abrir o novo documento
      var documento = DocumentApp.openById(novoDocumento.getId());

      // Pega corpo do documento
      var corpo = documento.getBody();

      // Substitui os marcadores pelos dados // Os marcadores são opcionais e de sua escolha, aqui usei "<<" e ">>"
      corpo.replaceText('<<nome>>', nomeAluno);
      corpo.replaceText('<<matricula>>', matriculaAluno);

      // Salva as alterações
      documento.saveAndClose();

      // Pega o arquivo do documento atualizado
      var arquivoAtualizado = DriveApp.getFileById(novoDocumento.getId());

      // Converte e salva em PDF
      var pdfBlob = arquivoAtualizado.getAs('application/pdf');
      pastaCerts.createFile(pdfBlob);
    }

    Logger.log('Certificados criados e salvos em PDF com sucesso na pasta "Certificados"!');
  } catch (erro) {
    Logger.log('Ocorreu um erro: ' + erro);
  }
}