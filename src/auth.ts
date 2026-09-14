export type AccessRole = "Administrador" | "Editor" | "Visualização";
export type Account = { id: string; name: string; email: string; role: AccessRole; salt: string; passwordHash: string };

const USERS_KEY = "roadmap.accounts.v1";
const SESSION_KEY = "roadmap.session.v1";

export function loadAccounts(): Account[] {
  try {
    const value = JSON.parse(localStorage.getItem(USERS_KEY) || "[]");
    return Array.isArray(value) ? value : [];
  } catch { return []; }
}

export function saveAccounts(accounts: Account[]) {
  localStorage.setItem(USERS_KEY, JSON.stringify(accounts));
}

export function sessionId() { return sessionStorage.getItem(SESSION_KEY); }
export function setSession(id: string | null) {
  if (id) sessionStorage.setItem(SESSION_KEY, id);
  else sessionStorage.removeItem(SESSION_KEY);
}

function hex(bytes: Uint8Array) { return Array.from(bytes, byte => byte.toString(16).padStart(2, "0")).join(""); }
export async function hashPassword(password: string, salt: string) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits({ name: "PBKDF2", salt: Uint8Array.from(salt.match(/../g)!.map(v => parseInt(v, 16))), iterations: 210000, hash: "SHA-256" }, key, 256);
  return hex(new Uint8Array(bits));
}

export async function makeAccount(name: string, email: string, password: string, role: AccessRole): Promise<Account> {
  const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
  return { id: crypto.randomUUID(), name: name.trim(), email: email.trim().toLowerCase(), role, salt, passwordHash: await hashPassword(password, salt) };
}
