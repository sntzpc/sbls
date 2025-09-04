// api.js — komunikasi dengan GAS Web App (CORS-safe via text/plain)

import { getConfig, setConfig, SHEET_HEADERS, toRow, fromRow, getAll, setAll } from './storage.js';

// SET URL WEB APP GAS ANDA DI SINI:
export const GAS_URL = 'https://script.google.com/macros/s/AKfycbxoiuCgvnY3X2pYwiD1tA1Lk5YOrg7_wxzTgQuFm3qwM2R7x36bOYSpQqFe33BSY1MJ/exec';

async function post(route, payload){
  const body = JSON.stringify({
    ...(payload||{}),
    sheetName: getConfig().sheetName,
    spreadsheetId: getConfig().spreadsheetId
  });
  const res = await fetch(`${GAS_URL}?route=${encodeURIComponent(route)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    mode: 'cors',
    body
  });
  if(!res.ok) throw new Error(`GAS ${route} HTTP ${res.status}`);
  return res.json();
}

// ===== API: INIT / HEADERS / PULL / PUSH
export async function apiInit(){
  const r = await post('init', { headers: SHEET_HEADERS });
  if(r?.spreadsheetId) setConfig({ spreadsheetId: r.spreadsheetId });
  return r;
}

export async function apiHeaders(){
  return post('headers', { headers: SHEET_HEADERS });
}

export async function apiPullOverwrite(){
  const r = await post('pull', {});
  const rows = (r?.rows||[]).map(fromRow);
  setAll(rows);
  return { ...r, count: rows.length };
}

export async function apiPushUpsert(){
  const all = getAll().map(toRow);
  return post('push', { rows: all });
}
