/* Alitaire Mail Agent – Office.js taskpane logica */

'use strict';

const SERVER = 'https://127.0.0.1:5757';

let huidigeMail = null;
let instellingenZichtbaar = false;

// ── Initialisatie ─────────────────────────────────────────────────────────────

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    controleerServer();
    laadMailInfo();
  }
});

// ── Server-controle ───────────────────────────────────────────────────────────

function controleerServer() {
  fetch(SERVER + '/api/ping', { method: 'GET' })
    .then(function(r) {
      if (!r.ok) throw new Error('Server antwoordt niet correct');
      toonEl('serverFout', false);
    })
    .catch(function() {
      toonEl('serverFout', true);
    });
}

// ── Mail-info laden ───────────────────────────────────────────────────────────

function laadMailInfo() {
  var item = Office.context.mailbox.item;
  if (!item) return;

  var onderwerp = item.subject || '(geen onderwerp)';
  var afzender  = '';
  var email     = '';

  if (item.from) {
    afzender = item.from.displayName  || '';
    email    = item.from.emailAddress || '';
  }

  item.body.getAsync(Office.CoercionType.Text, function(result) {
    var tekst = (result.status === Office.AsyncResultStatus.Succeeded)
      ? result.value
      : '';

    huidigeMail = {
      onderwerp:      onderwerp,
      afzender:       afzender,
      afzender_email: email,
      tekst:          tekst,
    };

    document.getElementById('mailOnderwerp').textContent = onderwerp;
    document.getElementById('mailAfzender').textContent  =
      afzender ? (afzender + (email ? ' <' + email + '>' : '')) : '(onbekend)';
    toonEl('mailInfo', true);
  });
}

// ── Reactie genereren ─────────────────────────────────────────────────────────

function genereerReactie() {
  if (!huidigeMail) {
    toonFout('Geen e-mail geladen. Sluit het paneel en open het opnieuw.');
    return;
  }

  zet('btnGenereer', 'disabled', true);
  toonEl('spinner', true);
  document.getElementById('btnTekst').textContent = 'Bezig met genereren...';
  toonEl('foutBlok', false);
  toonEl('resultaatBlok', false);

  fetch(SERVER + '/api/verwerk', {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(huidigeMail),
  })
    .then(function(r) {
      if (!r.ok) return r.json().then(function(d) { throw new Error(d.fout || r.statusText); });
      return r.json();
    })
    .then(function(data) {
      if (data.fout) throw new Error(data.fout);
      toonResultaat(data);
    })
    .catch(function(err) {
      var bericht = err.message || String(err);
      if (bericht.includes('Failed to fetch') || bericht.includes('NetworkError')) {
        bericht = 'Lokale server niet bereikbaar. Start start.bat op de laptop.';
        toonEl('serverFout', true);
      }
      toonFout(bericht);
    })
    .finally(function() {
      zet('btnGenereer', 'disabled', false);
      toonEl('spinner', false);
      document.getElementById('btnTekst').textContent = 'Reactie genereren';
    });
}

// ── Resultaat weergeven ───────────────────────────────────────────────────────

function toonResultaat(data) {
  var werf     = data.werf_naam       || '';
  var isNieuw  = data.is_nieuwe_werf;
  var antwoord = data.antwoord        || '';
  var pad      = data.opgeslagen_pad  || '';

  var badge = document.getElementById('werfBadge');
  badge.innerHTML = '📁 ' + escHtml(werf)
    + (isNieuw ? ' <span class="nieuw-badge">NIEUW</span>' : '');

  if (pad) {
    document.getElementById('opslaanPadTekst').textContent = pad;
    toonEl('opslaanPad', true);
  } else {
    toonEl('opslaanPad', false);
  }

  document.getElementById('antwoordArea').value = antwoord;
  toonEl('resultaatBlok', true);
}

// ── Acties ────────────────────────────────────────────────────────────────────

function voegIn() {
  var tekst = document.getElementById('antwoordArea').value;
  if (!tekst) return;

  var item = Office.context.mailbox.item;

  if (item.displayReplyFormAsync) {
    item.displayReplyFormAsync(
      { htmlBody: '<p>' + tekst.replace(/\n/g, '<br>') + '</p>' },
      function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          item.displayReplyForm({ htmlBody: '<p>' + tekst.replace(/\n/g, '<br>') + '</p>' });
        }
      }
    );
  } else {
    item.displayReplyForm({ htmlBody: '<p>' + tekst.replace(/\n/g, '<br>') + '</p>' });
  }
}

function kopieer() {
  var tekst = document.getElementById('antwoordArea').value;
  if (!tekst) return;

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tekst).then(function() {
      var knop = document.querySelector('.btn-kopieer');
      var orig = knop.textContent;
      knop.textContent = '✓ Gekopieerd';
      setTimeout(function() { knop.textContent = orig; }, 1500);
    });
  } else {
    var ta = document.getElementById('antwoordArea');
    ta.select();
    document.execCommand('copy');
  }
}

function opnieuw() {
  toonEl('resultaatBlok', false);
  genereerReactie();
}

// ── Instellingen ──────────────────────────────────────────────────────────────

function toggleInstellingen() {
  instellingenZichtbaar = !instellingenZichtbaar;
  toonEl('instellingenBlok', instellingenZichtbaar);
  if (instellingenZichtbaar) laadInstellingen();
}

function laadInstellingen() {
  fetch(SERVER + '/api/config')
    .then(function(r) { return r.json(); })
    .then(function(data) {
      document.getElementById('rootMap').value  = data.root_map   || '';
      document.getElementById('stijlMap').value = data.stijl_map  || '';
      if (data.heeft_api_sleutel) {
        document.getElementById('apiSleutel').placeholder = '(al ingesteld)';
      }
    })
    .catch(function() {});
}

function slaInstellingenOp() {
  var payload = {
    root_map:  document.getElementById('rootMap').value.trim(),
    stijl_map: document.getElementById('stijlMap').value.trim(),
  };
  var sleutel = document.getElementById('apiSleutel').value.trim();
  if (sleutel) payload.api_sleutel = sleutel;

  fetch(SERVER + '/api/config', {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(payload),
  })
    .then(function(r) { return r.json(); })
    .then(function(data) {
      if (data.ok) {
        var el = document.getElementById('meldingOk');
        el.style.display = 'block';
        setTimeout(function() { el.style.display = 'none'; }, 2000);
        document.getElementById('apiSleutel').value = '';
        document.getElementById('apiSleutel').placeholder = '(al ingesteld)';
      }
    })
    .catch(function(err) { toonFout('Config opslaan mislukt: ' + err.message); });
}

// ── Hulpfuncties ──────────────────────────────────────────────────────────────

function toonEl(id, zichtbaar) {
  var el = document.getElementById(id);
  if (el) el.style.display = zichtbaar ? '' : 'none';
}

function zet(id, attr, waarde) {
  var el = document.getElementById(id);
  if (!el) return;
  if (waarde === true)       el.setAttribute(attr, attr);
  else if (waarde === false) el.removeAttribute(attr);
  else                       el.setAttribute(attr, waarde);
}

function toonFout(bericht) {
  document.getElementById('foutTekst').textContent = bericht;
  toonEl('foutBlok', true);
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
