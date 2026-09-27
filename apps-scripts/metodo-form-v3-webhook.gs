/**
 * ENTRADA DE LEADS - FORM3 SHERAOS (v3)
 *
 * Recebe POST do form da /metodo-v2, salva TUDO na planilha com tracking completo,
 * e se faturamento >= 100k, encaminha pro CRM Sheraos (slug: sheraos-marketing)
 * com TODAS as tags, UTMs, cookies de tracking (fbp/fbc/fbclid/gclid/ctwa),
 * user_agent, page_url, event_id e notas ricas — o mesmo padrao dos outros LPs.
 *
 * Instalar:
 * 1. Abrir a planilha "ENTRADA DE LEADS - FORM3 SHERAOS"
 * 2. Extensoes -> Apps Script
 * 3. Colar TODO este arquivo, substituir Codigo.gs
 * 4. Salvar
 * 5. Rodar 1x a funcao initSheet (cria cabecalhos)
 * 6. Implantar -> Nova implantacao -> Tipo: App da web
 *    - Executar como: Eu mesmo
 *    - Quem tem acesso: Qualquer pessoa (mesmo anonimos)
 * 7. Copiar a URL da implantacao e passar pro dev pra atualizar em METODO_WEBHOOK no metodo-v2.html
 */

var SHEET_NAME = 'Pagina1' // ou 'Sheet1' se estiver em ingles
var CRM_WEBHOOK_URL = 'https://sheraos.com.br/crm/api/webhooks/sheets/sheraos-marketing'

// ─── Notificacao por email do novo lead ───────────────────────────
// Recebe UM email por lead que entra (mesmo os <100k, pra voce ver tudo)
var EMAIL_TO = 'agenciadouc@gmail.com'
var EMAIL_CC = ''  // ex: 'outro@email.com' — deixe vazio pra nao usar CC
var EMAIL_ENABLED = true  // false pra pausar notificacoes por email

// WhatsApp do atendente (numero de onde SAI a msg pro cliente).
// Usado no link wa.me — precisa ser o SEU numero (nao o do lead).
// Formato: DDI + DDD + numero, so digitos.
var WA_FROM = '554896838666'  // numero Sheraos (padrao). Muda se quiser outro.

// Template da mensagem que vai pre-preenchida no WhatsApp ao clicar o botao.
// {NOME} = primeiro nome do lead. Se vazio, usa "tudo bem?".
var WA_MSG_TEMPLATE = 'Olá {NOME}! Aqui é o João da Sheraos. Recebi seu contato pelo nosso site e queria entender melhor a sua empresa pra ver como podemos ajudar. Tem uns 10 minutinhos agora pra conversar?'

// Faturamentos que sao "acima de 100k" (vao pro CRM)
var HIGH_TIER = { '100k_200k': 1, '200k_300k': 1, 'acima_300k': 1 }

// Labels amigaveis pras notas do CRM
var FAT_LABEL = {
  'ate_50k': 'Ate R$ 50 mil',
  '50k_100k': 'R$ 50 a 100 mil',
  '100k_200k': 'R$ 100 a 200 mil',
  '200k_300k': 'R$ 200 a 300 mil',
  'acima_300k': 'Acima de R$ 300 mil'
}

var HEADERS = [
  'timestamp',           // 1
  'nome',                // 2
  'email',               // 3
  'whatsapp',            // 4
  'instagram',           // 5
  'faturamento',         // 6
  'categoria',           // 7
  'origem',              // 8
  'url',                 // 9
  'referrer',            // 10
  'user_agent',          // 11
  'utm_source',          // 12
  'utm_medium',          // 13
  'utm_campaign',        // 14
  'utm_content',         // 15
  'utm_term',            // 16
  'fbp',                 // 17
  'fbc',                 // 18
  'fbclid',              // 19
  'gclid',               // 20
  'ctwa_clid',           // 21
  'session_id',          // 22
  'ip',                  // 23
  'event_id',            // 24
  'crm_status',          // 25
  'crm_lead_id',         // 26
  'crm_response'         // 27
]

function doPost(e) {
  try {
    var data = {}
    try { data = JSON.parse(e.postData.contents) } catch (err) {
      data = e.parameter || {}
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0]

    // Garante cabecalhos na primeira linha
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS)
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f0f0f0')
      sh.setFrozenRows(1)
    }

    var faturamento = String(data.faturamento || '')
    var isHigh = !!HIGH_TIER[faturamento]
    var categoria = isHigh ? '>=100k' : '<100k'
    var eventId = data.event_id || _uuid_()

    var row = [
      new Date().toISOString(),
      data.nome || '',
      data.email || '',
      data.whatsapp || '',
      data.instagram || '',
      faturamento,
      categoria,
      data.origem || 'metodo-v2',
      data.url || '',
      data.referrer || '',
      data.user_agent || '',
      data.utm_source || '',
      data.utm_medium || '',
      data.utm_campaign || '',
      data.utm_content || '',
      data.utm_term || '',
      data.fbp || '',
      data.fbc || '',
      data.fbclid || '',
      data.gclid || '',
      data.ctwa_clid || '',
      data.session_id || '',
      data.ip || '',
      eventId,
      '', // crm_status (preenchido abaixo)
      '', // crm_lead_id
      ''  // crm_response
    ]

    sh.appendRow(row)
    var rowIndex = sh.getLastRow()
    var COL = { crm_status: 25, crm_lead_id: 26, crm_response: 27 }

    // ─── Notifica por email (nao bloqueia o response se der erro) ─
    if (EMAIL_ENABLED) {
      try { enviarEmailLead(data, faturamento, isHigh) }
      catch (err) { console.error('[email] falha:', err.message) }
    }

    // Se faturamento < 100k, NAO manda pro CRM. Marca como "pulado_baixo_faturamento".
    if (!isHigh) {
      sh.getRange(rowIndex, COL.crm_status).setValue('pulado_baixo_faturamento')
      return _json({ ok: true, categoria: categoria, sent_to_crm: false })
    }

    // ─── Monta tags ricas ─────────────────────────────
    var tags = ['LP Metodo V2', 'Alto Faturamento']
    if (FAT_LABEL[faturamento]) tags.push('Faturamento ' + FAT_LABEL[faturamento])
    if (data.utm_source) tags.push('utm:' + String(data.utm_source).toLowerCase())
    if (data.utm_campaign) tags.push('camp:' + String(data.utm_campaign).toLowerCase())
    if (data.gclid) tags.push('Origem Google Ads')
    if (data.fbclid) tags.push('Origem Meta Ads')
    if (!data.gclid && !data.fbclid && !data.utm_source) tags.push('Trafego Direto/Organico')

    // ─── Monta source_detail (aparece no card do CRM) ─
    var detailBits = []
    if (FAT_LABEL[faturamento]) detailBits.push('Faturamento: ' + FAT_LABEL[faturamento])
    if (data.instagram) detailBits.push('IG: ' + data.instagram)
    if (data.utm_campaign) detailBits.push('camp=' + data.utm_campaign)
    if (data.utm_content) detailBits.push('ad=' + data.utm_content)
    detailBits.push('LP: /metodo-v2')

    // ─── Notes ricas pro atendente ────────────────────
    var notesLines = []
    notesLines.push('Aplicou pelo LP /metodo-v2 em ' + _brDate_())
    notesLines.push('Faturamento declarado: ' + (FAT_LABEL[faturamento] || faturamento))
    if (data.instagram) notesLines.push('Instagram: ' + data.instagram)
    if (data.utm_source || data.utm_medium || data.utm_campaign) {
      notesLines.push('Origem: ' +
        [data.utm_source, data.utm_medium, data.utm_campaign].filter(function(x){return x}).join(' / '))
    }
    if (data.referrer) notesLines.push('Referrer: ' + data.referrer)
    if (data.url) notesLines.push('URL: ' + data.url)

    // ─── Payload completo pro CRM ─────────────────────
    // Atencao: nao mandar campos "extras" que virem no `notes` — a rota do CRM
    // concatena TODOS os campos nao-conhecidos como custom fields nas notas.
    // Entao mantemos so `notes` (rico) e campos que a rota reconhece explicitamente.
    var payload = {
      name: data.nome || '',
      phone: data.whatsapp || '',
      email: data.email || '',
      instagram: data.instagram || '',
      source: 'metodo-v2',
      source_detail: detailBits.join(' | ').substring(0, 500),
      tags: tags,
      notes: notesLines.join('\n'),

      // UTMs completos (todos reconhecidos pela rota)
      utm_source: data.utm_source || '',
      utm_medium: data.utm_medium || '',
      utm_campaign: data.utm_campaign || '',
      utm_content: data.utm_content || '',
      utm_term: data.utm_term || '',

      // Click IDs (CAPI + Google Ads matching)
      fbp: data.fbp || '',
      fbc: data.fbc || '',
      fbclid: data.fbclid || '',
      gclid: data.gclid || '',
      ctwa_clid: data.ctwa_clid || '',

      // Contexto do request pro Meta CAPI
      user_agent: data.user_agent || '',
      page_url: data.url || '',
      event_id: eventId
    }

    var crmResp
    try {
      var httpResp = UrlFetchApp.fetch(CRM_WEBHOOK_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        followRedirects: true
      })
      var code = httpResp.getResponseCode()
      var text = httpResp.getContentText()
      var respJson = null
      try { respJson = JSON.parse(text) } catch (err) {}
      var leadId = (respJson && (respJson.leadId || respJson.lead_id)) || ''

      if (code >= 200 && code < 300) {
        sh.getRange(rowIndex, COL.crm_status).setValue('ok_' + code)
        sh.getRange(rowIndex, COL.crm_lead_id).setValue(leadId)
        sh.getRange(rowIndex, COL.crm_response).setValue(text.substring(0, 400))
        crmResp = { ok: true, code: code, leadId: leadId }
      } else {
        sh.getRange(rowIndex, COL.crm_status).setValue('erro_' + code)
        sh.getRange(rowIndex, COL.crm_response).setValue(text.substring(0, 400))
        crmResp = { ok: false, code: code, body: text.substring(0, 200) }
      }
    } catch (err) {
      sh.getRange(rowIndex, COL.crm_status).setValue('excecao')
      sh.getRange(rowIndex, COL.crm_response).setValue(String(err).substring(0, 400))
      crmResp = { ok: false, error: String(err) }
    }

    return _json({ ok: true, categoria: categoria, sent_to_crm: true, crm: crmResp })
  } catch (err) {
    return _json({ ok: false, error: String(err) })
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}

function _uuid_() {
  // event_id estavel pra dedup no Meta CAPI. Formato UUIDv4-ish.
  var hex = '0123456789abcdef'
  var s = ''
  for (var i = 0; i < 32; i++) {
    var r = Math.floor(Math.random() * 16)
    if (i === 12) r = 4
    if (i === 16) r = (r & 0x3) | 0x8
    s += hex[r]
    if (i === 7 || i === 11 || i === 15 || i === 19) s += '-'
  }
  return s
}

function _brDate_() {
  var d = new Date()
  return Utilities.formatDate(d, 'America/Sao_Paulo', "dd/MM/yyyy 'as' HH:mm")
}

// ─── EMAIL DE NOTIFICACAO ───────────────────────────────────────
// Manda email pro EMAIL_TO com dados do lead + botao WhatsApp
// pre-preenchido usando WA_MSG_TEMPLATE (com {NOME} substituido).
function enviarEmailLead(data, faturamento, isHigh) {
  var nomeCliente = String(data.nome || '').trim()
  var primeiroNome = nomeCliente.split(/\s+/)[0] || 'tudo bem?'
  var telDigits = String(data.whatsapp || '').replace(/\D/g, '')
  var telFormatado = _formatarTelefone_(telDigits)

  var msg = WA_MSG_TEMPLATE.replace(/\{NOME\}/g, primeiroNome)
  var waLink = telDigits ? 'https://wa.me/' + telDigits + '?text=' + encodeURIComponent(msg) : ''
  var waFromLink = telDigits ? 'https://wa.me/' + WA_FROM + '?text=' + encodeURIComponent('Cliente novo: ' + nomeCliente + ' (' + telFormatado + ')') : ''

  var tierBadge = isHigh ? 'ALTO FATURAMENTO' : 'PADRAO'
  var tierColor = isHigh ? '#10B981' : '#8b5cf6'
  var fatLabel = FAT_LABEL[faturamento] || faturamento || '(nao informado)'

  var origemLinhas = []
  if (data.utm_source) origemLinhas.push(['Fonte', data.utm_source])
  if (data.utm_medium) origemLinhas.push(['Midia', data.utm_medium])
  if (data.utm_campaign) origemLinhas.push(['Campanha', data.utm_campaign])
  if (data.utm_content) origemLinhas.push(['Anuncio', data.utm_content])
  if (data.gclid) origemLinhas.push(['Google Ads', 'gclid presente'])
  if (data.fbclid) origemLinhas.push(['Meta Ads', 'fbclid presente'])
  if (!origemLinhas.length) origemLinhas.push(['Origem', 'Trafego direto / organico'])
  if (data.referrer) origemLinhas.push(['Referrer', data.referrer])

  var clienteLinhas = [
    ['Nome', nomeCliente || '(sem nome)'],
    ['WhatsApp', telFormatado || '(nao informado)'],
    ['Faturamento', fatLabel]
  ]
  if (data.email) clienteLinhas.push(['E-mail', data.email])
  if (data.instagram) clienteLinhas.push(['Instagram', data.instagram])

  var html = _buildEmailHtml_({
    tierBadge: tierBadge,
    tierColor: tierColor,
    primeiroNome: primeiroNome,
    nomeCliente: nomeCliente,
    telFormatado: telFormatado,
    origemLinhas: origemLinhas,
    clienteLinhas: clienteLinhas,
    waLink: waLink,
    msgPreview: msg
  })

  var plain = 'NOVO LEAD SHERAOS - ' + tierBadge + '\n\n' +
    'Nome: ' + nomeCliente + '\n' +
    'WhatsApp: ' + telFormatado + '\n' +
    'Faturamento: ' + fatLabel + '\n' +
    (data.instagram ? 'Instagram: ' + data.instagram + '\n' : '') +
    (waLink ? '\nAbrir WhatsApp com mensagem: ' + waLink + '\n' : '')

  var assunto = '[Sheraos] Lead ' + (isHigh ? '100k+' : '') + ' - ' + (nomeCliente || 'sem nome') + (telFormatado ? ' (' + telFormatado + ')' : '')

  var opts = {
    htmlBody: html,
    name: 'Leads Sheraos',
    replyTo: EMAIL_TO
  }
  if (EMAIL_CC) opts.cc = EMAIL_CC

  MailApp.sendEmail(EMAIL_TO, assunto, plain, opts)
}

function _formatarTelefone_(digits) {
  if (!digits) return ''
  var d = String(digits).replace(/\D/g, '')
  // Assume DDI 55 se vier com 12-13 digitos
  if (d.length === 13 && d.substring(0,2) === '55') d = d.substring(2)
  if (d.length === 12 && d.substring(0,2) === '55') d = d.substring(2)
  if (d.length === 11) return '(' + d.substring(0,2) + ') ' + d.substring(2,7) + '-' + d.substring(7)
  if (d.length === 10) return '(' + d.substring(0,2) + ') ' + d.substring(2,6) + '-' + d.substring(6)
  return digits
}

function _escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function _buildEmailHtml_(ctx) {
  function section(title, items, accent) {
    if (!items || !items.length) return ''
    var rows = items.map(function(kv) {
      return '<tr>' +
        '<td style="padding:9px 0;border-bottom:1px solid #E9E7F0;color:#6B6580;font-size:13px;font-weight:500;width:38%;vertical-align:top">' + _escapeHtml_(kv[0]) + '</td>' +
        '<td style="padding:9px 0;border-bottom:1px solid #E9E7F0;color:#1A1330;font-size:14px;font-weight:600;vertical-align:top">' + _escapeHtml_(kv[1]) + '</td>' +
      '</tr>'
    }).join('')
    return '<div style="margin-bottom:24px">' +
      '<h2 style="font-size:11px;text-transform:uppercase;letter-spacing:0.14em;color:' + accent + ';font-weight:800;margin:0 0 10px;padding-bottom:8px;border-bottom:2px solid ' + accent + '">' + _escapeHtml_(title) + '</h2>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse">' + rows + '</table>' +
    '</div>'
  }

  var waButton = ctx.waLink
    ? '<a href="' + _escapeHtml_(ctx.waLink) + '" target="_blank" style="display:inline-block;background:#25D366;color:#fff;padding:14px 24px;border-radius:10px;text-decoration:none;font-weight:700;font-size:15px;box-shadow:0 4px 14px rgba(37,211,102,0.35);letter-spacing:.01em">Abrir WhatsApp com ' + _escapeHtml_(ctx.primeiroNome) + '</a>'
    : '<div style="color:#9B96B0;font-size:13px">(sem whatsapp valido)</div>'

  var msgPreview = ctx.waLink
    ? '<div style="margin-top:14px;background:#F7F5FC;border-left:4px solid #6366F1;padding:12px 16px;border-radius:6px">' +
        '<div style="font-size:11px;text-transform:uppercase;letter-spacing:0.1em;color:#6B6580;font-weight:700;margin-bottom:6px">Mensagem pre-preenchida</div>' +
        '<div style="font-size:13px;color:#2A2119;line-height:1.5">' + _escapeHtml_(ctx.msgPreview) + '</div>' +
      '</div>'
    : ''

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:#EEEBF6;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;line-height:1.5;color:#1A1330">' +
    '<div style="padding:32px 16px">' +
    '<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" style="max-width:620px;width:100%;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 6px 30px rgba(20,15,40,0.10);border-collapse:collapse">' +

    // Header
    '<tr><td style="background:linear-gradient(135deg,#6366F1 0%,#4f46e5 60%,#3730A3 100%);padding:30px 32px;color:#fff">' +
      '<div style="display:inline-block;background:' + ctx.tierColor + ';padding:5px 12px;border-radius:20px;font-size:11px;font-weight:800;letter-spacing:0.14em;text-transform:uppercase;margin-bottom:12px">' + _escapeHtml_(ctx.tierBadge) + '</div>' +
      '<h1 style="font-size:22px;margin:0;font-weight:800;letter-spacing:-0.015em">Novo lead na Sheraos</h1>' +
      '<div style="font-size:14px;color:rgba(255,255,255,0.85);margin-top:6px">' + _escapeHtml_(ctx.nomeCliente || '(sem nome)') + (ctx.telFormatado ? ' &middot; ' + _escapeHtml_(ctx.telFormatado) : '') + '</div>' +
    '</td></tr>' +

    // CTA WhatsApp (topo pra ser 1o click)
    '<tr><td style="padding:28px 32px 8px;text-align:center">' +
      waButton +
      msgPreview +
    '</td></tr>' +

    // Sections
    '<tr><td style="padding:24px 32px 20px">' +
      section('Dados do cliente', ctx.clienteLinhas, '#4f46e5') +
      section('Origem do lead', ctx.origemLinhas, '#8b5cf6') +
    '</td></tr>' +

    // Footer
    '<tr><td style="background:#F7F5FC;padding:16px 32px;font-size:11.5px;color:#8A82A5;text-align:center;border-top:1px solid #E9E7F0">' +
      'Notificacao automatica do form <strong style="color:#4f46e5">/metodo-v2</strong>. Se faturamento >= 100k, o lead ja esta no CRM Sheraos.' +
    '</td></tr>' +

    '</table></div></body></html>'
}

// Util: rodar 1x pra testar so o email (sem gravar planilha nem CRM)
function testarEmailAgora() {
  enviarEmailLead({
    nome: 'Joao Teste da Silva',
    whatsapp: '48999998888',
    instagram: '@empresa_teste',
    faturamento: '100k_200k',
    email: 'teste@teste.com',
    utm_source: 'facebook',
    utm_medium: 'cpc',
    utm_campaign: 'metodo_v2_lead',
    utm_content: 'anuncio_v3'
  }, '100k_200k', true)
  Logger.log('Email enviado pra ' + EMAIL_TO + '. Confere sua caixa.')
}

// Util: rodar 1x manualmente pra criar cabecalhos vazios
function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0]
  sh.clear()
  sh.appendRow(HEADERS)
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f0f0f0')
  sh.setFrozenRows(1)
}

// Util: testar envio pro CRM manualmente (rodar 1x aqui pra confirmar integracao)
function testEnvioCRM() {
  var e = {
    postData: {
      contents: JSON.stringify({
        nome: 'Teste Lead Metodo V2',
        email: 'teste@teste.com',
        whatsapp: '48999998888',
        instagram: '@empresa_teste',
        faturamento: '100k_200k',
        origem: 'metodo-v2',
        url: 'https://sheraos.com.br/metodo-v2?utm_source=teste&utm_medium=manual',
        referrer: 'https://www.google.com',
        user_agent: 'Mozilla/5.0 Test',
        utm_source: 'teste',
        utm_medium: 'manual',
        utm_campaign: 'teste-integracao-crm',
        utm_content: 'anuncio-x',
        fbp: 'fb.1.1700000000000.123456789',
        fbclid: 'testFBCLID',
        session_id: 'sid_test123'
      })
    }
  }
  var r = doPost(e)
  Logger.log(r.getContent())
}
