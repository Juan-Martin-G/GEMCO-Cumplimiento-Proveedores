/*
 * Login por correo + código de 6 dígitos enviado con EmailJS.
 * Es una barrera básica (todo se valida en el navegador): evita accesos casuales
 * e indexación, no protege frente a alguien decidido. Ver README.
 */
'use strict';

const ACCESO = {
  emailjs: {
    publicKey: 'WGQ-1DDAFGlYSy2Qv',
    serviceId: 'service_l9u0laa',
    templateId: 'template_buuarj5', // variables: {{to_email}} y {{codigo}}
  },
  nombrePanel: 'Cumplimiento de proveedores GEMCO', // variable opcional {{panel}} para la plantilla
  timeoutEnvioMs: 15000,
  esperaReenvioSeg: 45,
  vigenciaCodigoMs: 15 * 60 * 1000,
  autorizados: [
    'juanmartin.brante@gemco.cl',
    'cristobal.munoz@aspencapital.cl',
    'pedro.delrio@gemco.cl',
    'diego.benavides@gemco.cl',
    'guillermo.pizarro@sassa.cl',
    'cristian.valenzuela@aspencapital.cl',
  ],
};

// Claves propias de este sitio (EERR comparte el dominio github.io).
const K = {
  auth: 'cumpl_auth', authEmail: 'cumpl_auth_email',
  email: 'cumpl_pending_email', codigo: 'cumpl_pending_code', vence: 'cumpl_pending_exp',
};

const $ = id => document.getElementById(id);
const ss = {
  get: k => { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
  set: (k, v) => { try { sessionStorage.setItem(k, v); } catch (e) { /* sin almacenamiento */ } },
  del: (...ks) => { try { ks.forEach(k => sessionStorage.removeItem(k)); } catch (e) { /* sin almacenamiento */ } },
};

try {
  const t = localStorage.getItem('tema');
  if (t === 'light' || t === 'dark') document.documentElement.dataset.theme = t;
} catch (e) { /* sin almacenamiento */ }

if (ss.get(K.auth) === 'true') location.replace('panel.html');

const normEmail = e => String(e || '').trim().toLowerCase();
const esAutorizado = e => ACCESO.autorizados.includes(normEmail(e));
function generarCodigo() {
  const n = new Uint32Array(1);
  crypto.getRandomValues(n);
  return String(100000 + (n[0] % 900000));
}

function mensaje(el, texto, tipo) {
  el.textContent = texto;
  el.className = `login-msg ${tipo}`;
  el.hidden = false;
}
const ocultar = el => { el.hidden = true; };

function irAPaso1() {
  $('paso2').hidden = true;
  $('paso1').hidden = false;
  ocultar($('msg2'));
  $('in-email').focus();
}
function irAPaso2(email) {
  $('email-eco').textContent = email;
  $('paso1').hidden = true;
  $('paso2').hidden = false;
  $('in-codigo').value = '';
  ocultar($('msg1'));
  $('in-codigo').focus();
}

// Traduce el error de EmailJS / red a un motivo legible.
function motivo(err) {
  if (err && typeof err.status === 'number') {
    return err.status === 0 ? 'sin conexión con api.emailjs.com (red o firewall)'
      : `EmailJS respondió ${err.status}${err.text ? ` (${err.text})` : ''}`;
  }
  if (err instanceof TypeError) return 'sin conexión con api.emailjs.com (red o firewall)';
  return (err && err.message) || 'error desconocido';
}

// emailjs.send con timeout: si api.emailjs.com no responde, la promesa del SDK
// queda pendiente para siempre; el timeout garantiza una salida.
async function enviar(email, codigo) {
  if (typeof emailjs === 'undefined') throw new Error('no se cargó la librería de EmailJS (revisa la conexión o un bloqueador)');
  const seg = Math.round(ACCESO.timeoutEnvioMs / 1000);
  let timer;
  const limite = new Promise((_, rechazar) => {
    timer = setTimeout(() => rechazar(new Error(`api.emailjs.com no respondió en ${seg} s (posible bloqueo de red o firewall)`)), ACCESO.timeoutEnvioMs);
  });
  try {
    return await Promise.race([
      emailjs.send(ACCESO.emailjs.serviceId, ACCESO.emailjs.templateId, { to_email: email, codigo, panel: ACCESO.nombrePanel }),
      limite,
    ]);
  } catch (err) {
    console.error('EmailJS:', err);
    throw new Error(motivo(err));
  } finally {
    clearTimeout(timer);
  }
}
function guardarPendiente(email, codigo) {
  ss.set(K.email, email);
  ss.set(K.codigo, codigo);
  ss.set(K.vence, String(Date.now() + ACCESO.vigenciaCodigoMs));
}

async function enviarCodigo(e) {
  e.preventDefault();
  const msg1 = $('msg1');
  const btn = $('btn-enviar');
  if (btn.disabled) return;
  const email = normEmail($('in-email').value);
  if (!esAutorizado(email)) { mensaje(msg1, 'Correo no autorizado.', 'err'); return; }
  const codigo = generarCodigo();
  btn.disabled = true;
  btn.textContent = 'Enviando…';
  ocultar(msg1);
  try {
    await enviar(email, codigo);
    guardarPendiente(email, codigo);
    irAPaso2(email);
    mensaje($('msg2'), 'Código enviado a tu correo, revisa tu bandeja de entrada.', 'ok');
    iniciarEspera();
  } catch (err) {
    mensaje(msg1, `No se pudo enviar el código: ${err.message}`, 'err');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Enviar código';
  }
}

async function reenviarCodigo() {
  const email = ss.get(K.email);
  if (!email) { irAPaso1(); return; }
  const codigo = generarCodigo();
  const btn = $('btn-reenviar');
  btn.disabled = true;
  btn.textContent = 'Enviando…';
  try {
    await enviar(email, codigo);
    guardarPendiente(email, codigo);
    mensaje($('msg2'), 'Código reenviado a tu correo.', 'ok');
    iniciarEspera();
  } catch (err) {
    mensaje($('msg2'), `No se pudo reenviar el código: ${err.message}`, 'err');
    btn.disabled = false;
    btn.textContent = 'Reenviar código';
  }
}

let temporizador = null;
function iniciarEspera() {
  const btn = $('btn-reenviar');
  let restante = ACCESO.esperaReenvioSeg;
  btn.disabled = true;
  btn.textContent = `Reenviar código (${restante}s)`;
  clearInterval(temporizador);
  temporizador = setInterval(() => {
    restante -= 1;
    if (restante <= 0) {
      clearInterval(temporizador);
      btn.disabled = false;
      btn.textContent = 'Reenviar código';
    } else {
      btn.textContent = `Reenviar código (${restante}s)`;
    }
  }, 1000);
}

function verificarCodigo(e) {
  e.preventDefault();
  const msg2 = $('msg2');
  const ingresado = $('in-codigo').value.trim();
  const codigo = ss.get(K.codigo);
  const email = ss.get(K.email);
  const vence = Number(ss.get(K.vence) || 0);
  if (!codigo || !email) { irAPaso1(); mensaje($('msg1'), 'Solicita un nuevo código.', 'err'); return; }
  if (Date.now() > vence) {
    ss.del(K.email, K.codigo, K.vence);
    irAPaso1();
    mensaje($('msg1'), 'El código expiró. Solicita uno nuevo.', 'err');
    return;
  }
  if (ingresado !== codigo) { mensaje(msg2, 'Código incorrecto.', 'err'); return; }
  ss.set(K.auth, 'true');
  ss.set(K.authEmail, email);
  ss.del(K.email, K.codigo, K.vence);
  location.replace('panel.html');
}

document.addEventListener('DOMContentLoaded', () => {
  if (typeof emailjs !== 'undefined') emailjs.init({ publicKey: ACCESO.emailjs.publicKey });
  $('paso1').addEventListener('submit', enviarCodigo);
  $('paso2').addEventListener('submit', verificarCodigo);
  $('btn-volver').addEventListener('click', irAPaso1);
  $('btn-reenviar').addEventListener('click', reenviarCodigo);
  $('in-email').focus();
});
