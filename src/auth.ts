export type AccessRole = "Administrador" | "Editor" | "Visualização";
export type Account = { id: string; name: string; email: string; role: AccessRole; salt: string; passwordHash: string };

const USERS_KEY = "roadmap.accounts.v1";
const SESSION_KEY = "roadmap.session.v1";
const ADMIN_SEED_KEY = "roadmap.admin-seed.v1";
export const ADMIN_EMAIL = "vagnersmoraes@hotmail.com";
const initialAdmin: Account = {
  id: "default-admin",
  name: "Vagner Moraes",
  email: ADMIN_EMAIL,
  role: "Administrador",
  salt: "f56bc8c844fdcadab43c9110c3db29f0",
  passwordHash: "6d4bc138c662228218ad86afe47dbba5a1da0387c6242f35b68e1e16f425bdc6",
};

export function loadAccounts(): Account[] {
  try {
    const value = JSON.parse(localStorage.getItem(USERS_KEY) || "[]");
    const accounts: Account[] = Array.isArray(value) ? value : [];
    if (localStorage.getItem(ADMIN_SEED_KEY)) return accounts;
    const existing = accounts.find(account => account.email?.toLowerCase() === ADMIN_EMAIL);
    const next = existing
      ? accounts.map(account => account.id === existing.id ? { ...account, email: ADMIN_EMAIL, role: "Administrador" as const, salt: initialAdmin.salt, passwordHash: initialAdmin.passwordHash } : account)
      : [initialAdmin, ...accounts];
    localStorage.setItem(USERS_KEY, JSON.stringify(next));
    localStorage.setItem(ADMIN_SEED_KEY, "1");
    return next;
  } catch { return [initialAdmin]; }
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
