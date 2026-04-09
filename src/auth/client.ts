import { JWT, OAuth2Client } from 'google-auth-library';
import * as fs from 'fs/promises';
import { getKeysFilePath, generateCredentialsErrorMessage, OAuthCredentials } from './utils.js';

// --- Service Account detection ---

const SA_KEY_PATH = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_FILE;
const SA_SUBJECT = process.env.USER_GOOGLE_EMAIL;

/** True when the server should use service account auth instead of user OAuth. */
export function isServiceAccountMode(): boolean {
  return !!SA_KEY_PATH;
}

/**
 * Create a JWT client from a service account key file with user impersonation.
 * JWT extends OAuth2Client, so it's compatible everywhere OAuth2Client is used.
 */
export async function initializeServiceAccountClient(): Promise<JWT> {
  if (!SA_KEY_PATH) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY_FILE is not set');
  }
  if (!SA_SUBJECT) {
    throw new Error(
      'USER_GOOGLE_EMAIL is required when using service account auth (the Workspace user to impersonate)'
    );
  }

  const keyContent = await fs.readFile(SA_KEY_PATH, 'utf-8');
  const key = JSON.parse(keyContent);

  if (key.type !== 'service_account') {
    throw new Error(
      `Expected "type": "service_account" in ${SA_KEY_PATH}, got "${key.type}"`
    );
  }

  const client = new JWT({
    email: key.client_email,
    key: key.private_key,
    scopes: ['https://www.googleapis.com/auth/calendar'],
    subject: SA_SUBJECT,
  });

  // Eagerly authorize to surface config errors at startup
  await client.authorize();
  process.stderr.write(
    `Service account authenticated (impersonating ${SA_SUBJECT})\n`
  );
  return client;
}

// --- OAuth credentials (existing path) ---

async function loadCredentialsFromFile(): Promise<OAuthCredentials> {
  const keysContent = await fs.readFile(getKeysFilePath(), "utf-8");
  const keys = JSON.parse(keysContent);

  if (keys.installed) {
    // Standard OAuth credentials file format
    const { client_id, client_secret, redirect_uris } = keys.installed;
    return { client_id, client_secret, redirect_uris };
  } else if (keys.client_id && keys.client_secret) {
    // Direct format
    return {
      client_id: keys.client_id,
      client_secret: keys.client_secret,
      redirect_uris: keys.redirect_uris || ['http://localhost:3000/oauth2callback']
    };
  } else {
    throw new Error('Invalid credentials file format. Expected either "installed" object or direct client_id/client_secret fields.');
  }
}

async function loadCredentialsWithFallback(): Promise<OAuthCredentials> {
  // Load credentials from file (CLI param, env var, or default path)
  try {
    return await loadCredentialsFromFile();
  } catch (fileError) {
    // Generate helpful error message
    const errorMessage = generateCredentialsErrorMessage();
    throw new Error(`${errorMessage}\n\nOriginal error: ${fileError instanceof Error ? fileError.message : fileError}`);
  }
}

export async function initializeOAuth2Client(): Promise<OAuth2Client> {
  // Service account mode — JWT extends OAuth2Client so the return type is compatible
  if (isServiceAccountMode()) {
    return initializeServiceAccountClient();
  }

  // Always use real OAuth credentials - no mocking.
  // Unit tests should mock at the handler level, integration tests need real credentials.
  try {
    const credentials = await loadCredentialsWithFallback();

    // Use the first redirect URI as the default for the base client
    return new OAuth2Client({
      clientId: credentials.client_id,
      clientSecret: credentials.client_secret,
      redirectUri: credentials.redirect_uris[0],
    });
  } catch (error) {
    throw new Error(`Error loading OAuth keys: ${error instanceof Error ? error.message : error}`);
  }
}

export async function loadCredentials(): Promise<{ client_id: string; client_secret: string }> {
  try {
    const credentials = await loadCredentialsWithFallback();

    if (!credentials.client_id || !credentials.client_secret) {
        throw new Error('Client ID or Client Secret missing in credentials.');
    }
    return {
      client_id: credentials.client_id,
      client_secret: credentials.client_secret
    };
  } catch (error) {
    throw new Error(`Error loading credentials: ${error instanceof Error ? error.message : error}`);
  }
}