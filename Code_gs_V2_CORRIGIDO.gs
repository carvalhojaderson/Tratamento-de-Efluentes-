/**
 * ============================================
 * SISTEMA DE ANALISE DE EFLUENTES - CODE.GS
 * Versao Corrigida - Evita criacao excessiva de planilhas
 * ============================================
 */

// ============================================
// CONFIGURACAO GLOBAL
// ============================================
const CONFIG = {
  NOME_PLANILHA: 'Sistema_Analise_Efluentes',
  ID_PLANILHA: null,
  ABAS: ['Usuarios', 'PontosColeta', 'Resultados', 'Graficos', 'Tokens', 'Analises', 'Artigos']
};

// ============================================
// FUNCAO DE INICIALIZACAO (EXECUTAR UMA VEZ)
// ============================================
function inicializarSistema() {
  try {
    let planilha = obterPlanilhaExistente();

    if (!planilha) {
      planilha = criarNovaPlanilha();
      Logger.log('Nova planilha criada: ' + planilha.getUrl());
    } else {
      Logger.log('Planilha existente encontrada: ' + planilha.getUrl());
    }

    configurarAbas(planilha);

    return {
      sucesso: true,
      mensagem: 'Sistema inicializado com sucesso!',
      url: planilha.getUrl(),
      id: planilha.getId()
    };
  } catch (erro) {
    Logger.log('Erro na inicializacao: ' + erro.toString());
    return {
      sucesso: false,
      mensagem: 'Erro: ' + erro.toString()
    };
  }
}

// ============================================
// OBTEM PLANILHA EXISTENTE
// ============================================
function obterPlanilhaExistente() {
  try {
    const arquivos = DriveApp.getFilesByName(CONFIG.NOME_PLANILHA);

    if (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      if (arquivo.getMimeType() === MimeType.GOOGLE_SHEETS) {
        return SpreadsheetApp.openById(arquivo.getId());
      }
    }

    try {
      const planilhas = SpreadsheetApp.getActiveSpreadsheet();
      if (planilhas && planilhas.getName() === CONFIG.NOME_PLANILHA) {
        return planilhas;
      }
    } catch (e) {
      // Ignora erro se nao houver planilha ativa
    }

    return null;
  } catch (erro) {
    Logger.log('Erro ao buscar planilha existente: ' + erro.toString());
    return null;
  }
}

// ============================================
// CRIA NOVA PLANILHA (APENAS SE NECESSARIO)
// ============================================
function criarNovaPlanilha() {
  try {
    const planilha = SpreadsheetApp.create(CONFIG.NOME_PLANILHA);
    planilha.setShareableByEditors(false);

    const abaPadrao = planilha.getSheetByName('Sheet1');
    if (abaPadrao) {
      planilha.deleteSheet(abaPadrao);
    }

    return planilha;
  } catch (erro) {
    throw new Error('Erro ao criar planilha: ' + erro.toString());
  }
}

// ============================================
// CONFIGURA AS ABAS NECESSARIAS
// ============================================
function configurarAbas(planilha) {
  const abasExistentes = planilha.getSheets().map(s => s.getName());

  CONFIG.ABAS.forEach(nomeAba => {
    if (!abasExistentes.includes(nomeAba)) {
      const novaAba = planilha.insertSheet(nomeAba);
      configurarCabecalhosAba(novaAba, nomeAba);
      Logger.log('Aba criada: ' + nomeAba);
    }
  });
}

// ============================================
// CONFIGURA CABECALHOS DE CADA ABA
// ============================================
function configurarCabecalhosAba(aba, nomeAba) {
  const cabecalhos = {
    'Usuarios': ['ID', 'Nome', 'Email', 'SenhaHash', 'Tokens', 'Nivel', 'DataRegistro', 'Ativo'],
    'PontosColeta': ['ID', 'Nome', 'Descricao', 'Localizacao', 'DataCadastro', 'Ativo'],
    'Resultados': ['ID', 'Data', 'Hora', 'UsuarioID', 'UsuarioNome', 'PontoColeta', 'TipoPonto', 'Parametro', 'Leitura1', 'Leitura2', 'Leitura3', 'Media', 'Unidade', 'Referencia', 'Observacoes', 'DataRegistro'],
    'Graficos': ['Data', 'PontoColeta', 'TipoPonto', 'Parametro', 'Media', 'Unidade'],
    'Tokens': ['ID', 'UsuarioID', 'Quantidade', 'Tipo', 'Descricao', 'Data'],
    'Analises': ['ID', 'UsuarioID', 'Data', 'Hora', 'PontoColeta', 'TipoPonto', 'ParametrosJSON', 'Observacoes', 'DataRegistro'],
    'Artigos': ['ID', 'UsuarioID', 'Titulo', 'ConteudoJSON', 'DataCriacao', 'DataAtualizacao']
  };

  if (cabecalhos[nomeAba]) {
    aba.getRange(1, 1, 1, cabecalhos[nomeAba].length).setValues([cabecalhos[nomeAba]]);
    aba.getRange(1, 1, 1, cabecalhos[nomeAba].length)
      .setFontWeight('bold')
      .setBackground('#1a5fb4')
      .setFontColor('white');
    aba.setFrozenRows(1);
  }
}

// ============================================
// FUNCOES DO SISTEMA WEB
// ============================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistema de Analise de Efluentes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// AUTENTICACAO
// ============================================
function loginUsuario(email, senha) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) {
      return { sucesso: false, mensagem: 'Sistema nao inicializado. Execute inicializarSistema() primeiro.' };
    }

    const abaUsuarios = planilha.getSheetByName('Usuarios');
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][2] === email && dados[i][3] === hashSenha(senha) && dados[i][7] === true) {
        return {
          sucesso: true,
          usuario: {
            id: dados[i][0],
            nome: dados[i][1],
            email: dados[i][2],
            tokens: dados[i][4],
            nivel: dados[i][5]
          }
        };
      }
    }

    return { sucesso: false, mensagem: 'Email ou senha incorretos' };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro: ' + erro.toString() };
  }
}

function registrarUsuario(dados) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) {
      return { sucesso: false, mensagem: 'Sistema nao inicializado.' };
    }

    const abaUsuarios = planilha.getSheetByName('Usuarios');
    const dadosExistentes = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dadosExistentes.length; i++) {
      if (dadosExistentes[i][2] === dados.email) {
        return { sucesso: false, mensagem: 'Email ja cadastrado' };
      }
    }

    const novoId = 'USR_' + new Date().getTime();

    abaUsuarios.appendRow([
      novoId,
      dados.nome,
      dados.email,
      hashSenha(dados.senha),
      100,
      'Iniciante',
      new Date(),
      true
    ]);

    return { sucesso: true, mensagem: 'Usuario cadastrado com sucesso!' };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro: ' + erro.toString() };
  }
}

function hashSenha(senha) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha)
    .map(function(byte) {
      return (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0');
    })
    .join('');
}

// ============================================
// ANALISES
// ============================================
function salvarAnalise(dados) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) {
      return { sucesso: false, mensagem: 'Sistema nao inicializado.' };
    }

    const abaResultados = planilha.getSheetByName('Resultados');
    const abaAnalises = planilha.getSheetByName('Analises');
    const abaGraficos = planilha.getSheetByName('Graficos');

    const analiseId = 'ANL_' + new Date().getTime();
    const dataRegistro = new Date();

    dados.parametros.forEach(param => {
      abaResultados.appendRow([
        analiseId,
        dados.data,
        dados.hora,
        dados.usuarioId,
        '',
        dados.pontoColeta,
        dados.tipoPonto,
        param.nome,
        param.leitura1 || '',
        param.leitura2 || '',
        param.leitura3 || '',
        param.media,
        param.unidade,
        param.referencia,
        dados.observacoes || '',
        dataRegistro
      ]);

      abaGraficos.appendRow([
        dados.data,
        dados.pontoColeta,
        dados.tipoPonto,
        param.nome,
        parseFloat(param.media) || 0,
        param.unidade
      ]);
    });

    abaAnalises.appendRow([
      analiseId,
      dados.usuarioId,
      dados.data,
      dados.hora,
      dados.pontoColeta,
      dados.tipoPonto,
      JSON.stringify(dados.parametros),
      dados.observacoes || '',
      dataRegistro
    ]);

    const tokensGanhos = dados.parametros.length * 10;
    creditarTokens(dados.usuarioId, tokensGanhos, 'analise', 'Analise ' + analiseId);

    return {
      sucesso: true,
      mensagem: 'Analise salva com sucesso!',
      analiseId: analiseId,
      tokensGanhos: tokensGanhos
    };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro: ' + erro.toString() };
  }
}

function obterHistoricoAnalises(filtros) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) return [];

    const abaAnalises = planilha.getSheetByName('Analises');
    const dados = abaAnalises.getDataRange().getValues();
    const analises = [];

    for (let i = 1; i < dados.length; i++) {
      const analise = {
        id: dados[i][0],
        usuarioId: dados[i][1],
        data: dados[i][2],
        hora: dados[i][3],
        pontoColeta: dados[i][4],
        tipoPonto: dados[i][5],
        parametros: JSON.parse(dados[i][6] || '[]'),
        observacoes: dados[i][7]
      };

      let incluir = true;
      if (filtros.usuario && analise.usuarioId !== filtros.usuario) incluir = false;
      if (filtros.dataInicio && analise.data < filtros.dataInicio) incluir = false;
      if (filtros.dataFim && analise.data > filtros.dataFim) incluir = false;
      if (filtros.pontoColeta && analise.pontoColeta !== filtros.pontoColeta) incluir = false;
      if (filtros.tipoPonto && analise.tipoPonto !== filtros.tipoPonto) incluir = false;

      if (incluir) analises.push(analise);
    }

    return analises.reverse();
  } catch (erro) {
    Logger.log('Erro ao obter historico: ' + erro.toString());
    return [];
  }
}

// ============================================
// PONTOS DE COLETA
// ============================================
function salvarPontoColeta(dados) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) {
      return { sucesso: false, mensagem: 'Sistema nao inicializado.' };
    }

    const abaPontos = planilha.getSheetByName('PontosColeta');
    const novoId = 'PTC_' + new Date().getTime();

    abaPontos.appendRow([
      novoId,
      dados.nome,
      dados.descricao || '',
      dados.localizacao || '',
      new Date(),
      true
    ]);

    return { sucesso: true, mensagem: 'Ponto de coleta salvo!' };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro: ' + erro.toString() };
  }
}

function obterPontosColeta() {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) return [];

    const abaPontos = planilha.getSheetByName('PontosColeta');
    const dados = abaPontos.getDataRange().getValues();
    const pontos = [];

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][5] === true) {
        pontos.push({
          id: dados[i][0],
          nome: dados[i][1],
          descricao: dados[i][2],
          localizacao: dados[i][3]
        });
      }
    }

    return pontos;
  } catch (erro) {
    return [];
  }
}

// ============================================
// TOKENS
// ============================================
function creditarTokens(usuarioId, quantidade, tipo, descricao) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) return;

    const abaTokens = planilha.getSheetByName('Tokens');
    const novoId = 'TKN_' + new Date().getTime();

    abaTokens.appendRow([
      novoId,
      usuarioId,
      quantidade,
      tipo,
      descricao,
      new Date()
    ]);

    const abaUsuarios = planilha.getSheetByName('Usuarios');
    const dadosUsuarios = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dadosUsuarios.length; i++) {
      if (dadosUsuarios[i][0] === usuarioId) {
        const saldoAtual = dadosUsuarios[i][4] || 0;
        abaUsuarios.getRange(i + 1, 5).setValue(saldoAtual + quantidade);
        break;
      }
    }
  } catch (erro) {
    Logger.log('Erro ao creditar tokens: ' + erro.toString());
  }
}

function obterSaldoTokens(usuarioId) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) return 0;

    const abaUsuarios = planilha.getSheetByName('Usuarios');
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === usuarioId) {
        return dados[i][4] || 0;
      }
    }

    return 0;
  } catch (erro) {
    return 0;
  }
}

// ============================================
// GRAFICOS
// ============================================
function obterDadosGraficos() {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) {
      return { sucesso: false, mensagem: 'Sistema nao inicializado.' };
    }

    const abaGraficos = planilha.getSheetByName('Graficos');
    const dados = abaGraficos.getDataRange().getValues();

    const dadosPorParametro = {};

    for (let i = 1; i < dados.length; i++) {
      const parametro = dados[i][3];
      const tipoPonto = dados[i][2];
      const media = parseFloat(dados[i][4]) || 0;

      if (!dadosPorParametro[parametro]) {
        dadosPorParametro[parametro] = { entrada: 0, intermediario: 0, saida: 0, count: { entrada: 0, intermediario: 0, saida: 0 } };
      }

      const tipo = tipoPonto.toLowerCase();
      if (tipo === 'entrada') {
        dadosPorParametro[parametro].entrada += media;
        dadosPorParametro[parametro].count.entrada++;
      } else if (tipo === 'intermediario') {
        dadosPorParametro[parametro].intermediario += media;
        dadosPorParametro[parametro].count.intermediario++;
      } else if (tipo === 'saida') {
        dadosPorParametro[parametro].saida += media;
        dadosPorParametro[parametro].count.saida++;
      }
    }

    Object.keys(dadosPorParametro).forEach(param => {
      const d = dadosPorParametro[param];
      d.entrada = d.count.entrada > 0 ? d.entrada / d.count.entrada : 0;
      d.intermediario = d.count.intermediario > 0 ? d.intermediario / d.count.intermediario : 0;
      d.saida = d.count.saida > 0 ? d.saida / d.count.saida : 0;
    });

    return { sucesso: true, dados: dadosPorParametro };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro: ' + erro.toString() };
  }
}

// ============================================
// ARTIGOS
// ============================================
function salvarArtigo(dados) {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) {
      return { sucesso: false, mensagem: 'Sistema nao inicializado.' };
    }

    const abaArtigos = planilha.getSheetByName('Artigos');
    const novoId = 'ART_' + new Date().getTime();

    abaArtigos.appendRow([
      novoId,
      dados.usuarioId,
      dados.titulo,
      JSON.stringify(dados),
      new Date(),
      new Date()
    ]);

    creditarTokens(dados.usuarioId, 50, 'artigo', 'Artigo salvo: ' + dados.titulo);

    return { sucesso: true, mensagem: 'Artigo salvo!', tokensGanhos: 50 };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro: ' + erro.toString() };
  }
}

// ============================================
// FUNCOES DE TESTE
// ============================================
function testarConexao() {
  const planilha = obterPlanilhaExistente();
  if (planilha) {
    Logger.log('Conexao OK: ' + planilha.getName());
    return 'Conexao bem-sucedida com: ' + planilha.getName();
  } else {
    Logger.log('Nenhuma planilha encontrada');
    return 'Nenhuma planilha encontrada. Execute inicializarSistema() primeiro.';
  }
}

function limparDados() {
  try {
    const planilha = obterPlanilhaExistente();
    if (!planilha) return 'Planilha nao encontrada';

    CONFIG.ABAS.forEach(nomeAba => {
      const aba = planilha.getSheetByName(nomeAba);
      if (aba) {
        const ultimaLinha = aba.getLastRow();
        if (ultimaLinha > 1) {
          aba.deleteRows(2, ultimaLinha - 1);
        }
      }
    });

    return 'Dados limpos!';
  } catch (erro) {
    return 'Erro: ' + erro.toString();
  }
}
