// Prueba de Fichas de dilución en Edge sin interfaz (CDP), sin instalar nada (Node 22+).
// Uso: npm run servir   (en otra terminal)
//      node herramientas/prueba-navegador.mjs <carpeta-capturas> [light|dark]
import { spawn } from 'node:child_process';
import { mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const salida = process.argv[2] || 'capturas';
const modo = process.argv[3] || 'light';
const BASE = process.env.BASE || 'http://127.0.0.1:5507/';
const PUERTO = 9334;
mkdirSync(salida, { recursive: true });
const perfil = join(tmpdir(), 'edge-fichas-' + Date.now());
const edge = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  ['--headless=new', `--remote-debugging-port=${PUERTO}`, `--user-data-dir=${perfil}`, '--no-first-run', 'about:blank']);
const esperar = ms => new Promise(r => setTimeout(r, ms));

let lista = [];
for (let i = 0; i < 40; i++) {
  try { lista = await (await fetch(`http://127.0.0.1:${PUERTO}/json`)).json(); if (lista.find(t => t.type === 'page')) break; } catch { /* Edge aún no abre */ }
  await esperar(250);
}
const ws = new WebSocket(lista.find(t => t.type === 'page').webSocketDebuggerUrl);
await new Promise(r => { ws.onopen = r; });

let id = 0;
const pendientes = {};
const errores = [];
ws.onmessage = ev => {
  const m = JSON.parse(ev.data);
  if (m.id && pendientes[m.id]) { pendientes[m.id](m); delete pendientes[m.id]; }
  if (m.method === 'Runtime.exceptionThrown') errores.push(m.params.exceptionDetails.exception?.description || m.params.exceptionDetails.text);
  if (m.method === 'Runtime.consoleAPICalled' && ['error', 'warning'].includes(m.params.type)) errores.push(m.params.type + ': ' + m.params.args.map(a => a.value ?? a.description).join(' '));
  if (m.method === 'Log.entryAdded' && m.params.entry.level === 'error') errores.push('log: ' + m.params.entry.text + ' ' + (m.params.entry.url || ''));
};
const cdp = (method, params = {}) => new Promise(r => { const i = ++id; pendientes[i] = r; ws.send(JSON.stringify({ id: i, method, params })); });
const evaluar = async expr => (await cdp('Runtime.evaluate', { expression: expr, awaitPromise: true, returnByValue: true })).result?.result?.value;
const tam = (w, h, movil) => cdp('Emulation.setDeviceMetricsOverride', { width: w, height: h, deviceScaleFactor: 1, mobile: movil });
const ir = async url => { await cdp('Page.navigate', { url }); await esperar(1800); };
const buscar = async texto => {
  await evaluar(`(() => { const i = document.getElementById('buscar'); i.value = ${JSON.stringify(texto)}; i.dispatchEvent(new Event('input')); })()`);
  await esperar(300);
  return evaluar('document.getElementById("contador").textContent');
};
const captura = async (nombre, completa = true) => {
  let clip;
  if (completa) {
    const alto = await evaluar('document.documentElement.scrollHeight');
    const ancho = await evaluar('document.documentElement.clientWidth');
    clip = { x: 0, y: 0, width: ancho, height: Math.min(alto, 6000), scale: 1 };
  }
  const r = await cdp('Page.captureScreenshot', { format: 'png', captureBeyondViewport: completa, ...(clip ? { clip } : {}) });
  writeFileSync(join(salida, `${nombre}-${modo}.png`), Buffer.from(r.result.data, 'base64'));
};

await cdp('Runtime.enable'); await cdp('Log.enable'); await cdp('Page.enable');
await cdp('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: modo }] });

// --ESCRITORIO
await tam(1280, 900, false);
await ir(BASE);
console.log('Inicio:', await evaluar('document.getElementById("contador").textContent'));
console.log('Logos cargados:', await evaluar('[...document.querySelectorAll("header img")].map(i => i.alt.split(" ")[0] + "=" + (i.complete && i.naturalWidth > 0)).join(", ")'));
await captura('01-inicio');
console.log('Buscar "norad":', await buscar('norad'));
console.log('Buscar "vitalis":', await buscar('vitalis'));
await buscar('xyzw'); await captura('02-sin-resultados', false);
await buscar('');
await evaluar('document.querySelector("[data-filtro=periferica]").click()'); await esperar(300);
console.log('Filtro periférica:', await evaluar('document.getElementById("contador").textContent'));
await ir(BASE + '#/m/noradrenalina-4-mg'); await captura('03-ficha-noradrenalina');
await ir(BASE + '#/m/ceftriaxona-1-g/2'); await captura('04-ficha-ceftriaxona-roche');
console.log('Pestañas ceftriaxona:', await evaluar('[...document.querySelectorAll("[role=tab]")].map(a => a.textContent.trim()).join(" | ")'));
for (const g of ['soluciones', 'insulina', 'periferica', 'subcutanea', 'uso']) { await ir(BASE + '#/guia/' + g); await captura('05-guia-' + g); }

// --TODAS LAS FICHAS ABREN CON CONTENIDO
const ids = await evaluar('fetch("datos/medicamentos.json").then(r => r.json()).then(d => d.medicamentos.flatMap(m => m.presentaciones.map((_, j) => m.id + "/" + j)))');
const malas = [];
for (const x of ids) {
  await cdp('Page.navigate', { url: BASE + '#/m/' + x }); await esperar(60);
  if (!await evaluar('!!document.querySelector("article h1") && document.querySelectorAll("dd").length')) malas.push(x);
}
console.log(`Fichas revisadas: ${ids.length}, sin contenido: ${malas.join(', ') || 'ninguna'}`);

// --CELULAR
await tam(390, 844, true);
await ir(BASE + '#/'); await captura('06-celular-inicio', false);
await ir(BASE + '#/m/piperacilina-tazobactam-4-5-g'); await captura('07-celular-ficha');
console.log('Desborde horizontal en celular:', await evaluar('document.documentElement.scrollWidth > document.documentElement.clientWidth'));
console.log('Consola:', errores.length ? '\n  ' + errores.join('\n  ') : 'sin errores');

ws.close(); edge.kill(); await esperar(500);
try { rmSync(perfil, { recursive: true, force: true }); } catch { /* Edge aún lo tiene abierto */ }
