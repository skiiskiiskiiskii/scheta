// Хранение в браузере (IndexedDB): настройки, таблица, журнал, кэш распознаваний, незавершённая пачка.
// Всё лежит только в этом браузере на этом компьютере.
const DB_NAME = "scheta", DB_VER = 1;
let dbp = null;

function db() {
  if (!dbp) {
    dbp = new Promise((resolve, reject) => {
      const r = indexedDB.open(DB_NAME, DB_VER);
      r.onupgradeneeded = () => {
        const d = r.result;
        for (const s of ["kv", "cache", "files"]) if (!d.objectStoreNames.contains(s)) d.createObjectStore(s);
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }
  return dbp;
}

async function tx(store, mode, fn) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction(store, mode);
    const req = fn(t.objectStore(store));
    t.oncomplete = () => resolve(req?.result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error || new Error("запись в браузер не удалась"));
  });
}

const make = (store) => ({
  get: (k) => tx(store, "readonly", (s) => s.get(k)),
  set: (k, v) => tx(store, "readwrite", (s) => s.put(v, k)),
  del: (k) => tx(store, "readwrite", (s) => s.delete(k)),
  clear: () => tx(store, "readwrite", (s) => s.clear()),
});

export const kv = make("kv");
export const cache = make("cache");
export const files = make("files");

/** Просим браузер не удалять наши данные при нехватке места. */
export async function askPersistent() {
  try { return navigator.storage?.persist ? await navigator.storage.persist() : false; } catch { return false; }
}

export async function wipeAll() {
  await kv.clear(); await cache.clear(); await files.clear();
}

// ── таблица ──
// table = { rows: [...], keys: {key: "где"}, hashes: {sha: "где"}, journal: [...], exportedRows: n, name }
export const emptyTable = () => ({ rows: [], keys: {}, hashes: {}, journal: [], exportedRows: 0, created: new Date().toISOString() });
