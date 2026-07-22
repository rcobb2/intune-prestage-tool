import { resolve } from 'path';

const e = (k: string) => JSON.stringify(process.env[k] ?? '');

await Bun.build({
  entrypoints: ['./main.ts'],
  outdir: './client/',
  // Bun 1.2+ dropped the env-object API; use define to explicitly allowlist vars.
  env: 'disable',
  define: {
    'process.env.SKIP_ENTRA_AUTH':    e('SKIP_ENTRA_AUTH'),
    'process.env.AZURE_CLIENT_ID':    e('AZURE_CLIENT_ID'),
    'process.env.AZURE_AUTHORITY':    e('AZURE_AUTHORITY'),
    'process.env.CLIENT_HOSTNAME':    e('CLIENT_HOSTNAME'),
    'process.env.CLIENT_PORT':        e('CLIENT_PORT'),
    'process.env.THEME':              e('THEME'),
    'process.env.SERVER_API_HOSTNAME': e('SERVER_API_HOSTNAME'),
    'process.env.SERVER_API_PORT':    e('SERVER_API_PORT'),
  },
  target: 'browser',
  format: 'esm',
  sourcemap: 'none',
  splitting: false,
  minify: true,
});

console.log('Client build completed successfully.');

// TLS is only enabled when cert/key files are present (self-signed local/docker-compose dev).
// Behind a reverse proxy (e.g. Caddy in prod) no certs are provided and this serves plain HTTP.
const hasTls = await Bun.file("server.cert").exists() && await Bun.file("server.key").exists();

// CLIENT_BIND_HOST/PORT decouple the container's actual listen address from
// CLIENT_HOSTNAME/PORT (the public address baked into the browser bundle above),
// so a reverse proxy can front a different public host/port. Both fall back to the
// public values, so single-host/no-proxy setups are unaffected.
const server = Bun.serve({
  development: false,
  hostname: process.env.CLIENT_BIND_HOST || process.env.CLIENT_HOSTNAME || "localhost",
  port: process.env.CLIENT_BIND_PORT || process.env.CLIENT_PORT || 3000,
  ...(hasTls ? {
    tls: {
      key: Bun.file("server.key"),
      cert: Bun.file("server.cert"),
    },
  } : {}),
  async fetch(req) {
    const url = new URL(req.url);
    const path = url.pathname;

    // Serve the homepage
    if (path === "/" || path === "/index.html") {
      const file = Bun.file("/app/client/index.html");
      return new Response(file, {
        headers: { "Content-Type": "text/html" },
      });
    }

    // Serve static files from the client directory — resolve to prevent path traversal
    const filePath = resolve('/app/client', '.' + path);
    if (!filePath.startsWith('/app/client/')) {
      return new Response('Forbidden', { status: 403 });
    }
    const file = Bun.file(filePath);

    if (await file.exists()) {
      return new Response(file, {
        headers: { "Cache-Control": "no-store" },
      });
    }

    // Return 404 for everything else
    const notFoundFile = Bun.file("/app/404.html");
    return new Response(notFoundFile, {
      status: 404,
      headers: { "Content-Type": "text/html" },
    });
  },
});

console.log(`Bun version: ${Bun.version_with_sha}`);
console.log(`Client listening on ${server.url}`);
