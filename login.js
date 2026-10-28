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
  auth: 'gcp_auth', authEmail: 'gcp_auth_email',
  email: 'gcp_pending_email', codigo: 'gcp_pending_code', vence: 'gcp_pending_exp',
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

function enviar(email, codigo) {
  if (typeof emailjs === 'undefined') return Promise.reject(new Error('EmailJS no cargó'));
  return emailjs.send(ACCESO.emailjs.serviceId, ACCESO.emailjs.templateId, { to_email: email, codigo }, { publicKey: ACCESO.emailjs.publicKey });
}
function guardarPendiente(email, codigo) {
  ss.set(K.email, email);
  ss.set(K.codigo, codigo);
  ss.set(K.vence, String(Date.now() + ACCESO.vigenciaCodigoMs));
}

function enviarCodigo(e) {
  e.preventDefault();
  const msg1 = $('msg1');
  const email = normEmail($('in-email').value);
  if (!esAutorizado(email)) { mensaje(msg1, 'Correo no autorizado.', 'err'); return; }
  const codigo = generarCodigo();
  const btn = $('btn-enviar');
  btn.disabled = true;
  ocultar(msg1);
  enviar(email, codigo).then(() => {
    guardarPendiente(email, codigo);
    btn.disabled = false;
    irAPaso2(email);
    mensaje($('msg2'), 'Código enviado a tu correo, revisa tu bandeja de entrada.', 'ok');
    iniciarEspera();
  }).catch(() => {
    btn.disabled = false;
    mensaje(msg1, 'No se pudo enviar el código. Intenta nuevamente.', 'err');
  });
}

function reenviarCodigo() {
  const email = ss.get(K.email);
  if (!email) { irAPaso1(); return; }
  const codigo = generarCodigo();
  const btn = $('btn-reenviar');
  btn.disabled = true;
  enviar(email, codigo).then(() => {
    guardarPendiente(email, codigo);
    mensaje($('msg2'), 'Código reenviado a tu correo.', 'ok');
    iniciarEspera();
  }).catch(() => {
    mensaje($('msg2'), 'No se pudo reenviar el código. Intenta nuevamente.', 'err');
    btn.disabled = false;
  });
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
  $('paso1').addEventListener('submit', enviarCodigo);
  $('paso2').addEventListener('submit', verificarCodigo);
  $('btn-volver').addEventListener('click', irAPaso1);
  $('btn-reenviar').addEventListener('click', reenviarCodigo);
  $('in-email').focus();
});
