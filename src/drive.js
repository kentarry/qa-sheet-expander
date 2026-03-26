import { APPS_SCRIPT_URL } from './config.js';

export async function driveListFolder(folderId) {
  const r = await fetch(APPS_SCRIPT_URL + '?action=list&folderId=' + folderId);
  const d = await r.json();
  if (d.error) throw new Error(d.error);
  return d; // { folderName, folders:[], files:[] }
}

export async function driveDownloadFile(fileId) {
  const r = await fetch(APPS_SCRIPT_URL + '?action=download&fileId=' + fileId);
  const d = await r.json();
  if (d.error) throw new Error(d.error);
  const bin = atob(d.base64);
  const buf = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
  return { buf, name: d.fileName };
}
