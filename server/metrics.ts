import client from 'prom-client';

export const register = new client.Registry();
client.collectDefaultMetrics({ register });

export const httpRequestDuration = new client.Histogram({
  name: 'http_request_duration_seconds',
  help: 'Duration of HTTP requests in seconds',
  labelNames: ['method', 'route', 'status'] as const,
  buckets: [0.01, 0.05, 0.1, 0.3, 0.5, 1, 2, 5, 10],
  registers: [register],
});

export const httpRequestsTotal = new client.Counter({
  name: 'http_requests_total',
  help: 'Total number of HTTP requests',
  labelNames: ['method', 'route', 'status'] as const,
  registers: [register],
});

// Outbound calls to Microsoft Graph / GLPI / ClearPass are the app's external dependencies;
// tracked separately from inbound HTTP, and by target, so a slow/failing upstream
// is visible independent of how the UI itself is performing.
export const externalApiRequestDuration = new client.Histogram({
  name: 'external_api_request_duration_seconds',
  help: 'Duration of outbound requests to external APIs (Microsoft Graph, GLPI, ClearPass) in seconds',
  labelNames: ['target', 'method', 'status'] as const,
  buckets: [0.05, 0.1, 0.3, 0.5, 1, 2, 5, 10],
  registers: [register],
});

export const externalApiErrorsTotal = new client.Counter({
  name: 'external_api_errors_total',
  help: 'Total number of failed outbound requests to external APIs (Microsoft Graph, GLPI, ClearPass)',
  labelNames: ['target', 'method', 'status'] as const,
  registers: [register],
});

// Wraps a route handler to record request count/duration under a fixed route
// label (the route's static path pattern, e.g. "/api/devices/:search") so
// per-request values like serial numbers never become metric labels.
export function withMetrics(route: string, handler: (req: any) => Response | Promise<Response>) {
  return async (req: any): Promise<Response> => {
    const start = performance.now();
    let status = 500;
    try {
      const res = await handler(req);
      status = res.status;
      return res;
    } finally {
      const labels = { method: req.method, route, status: String(status) };
      httpRequestDuration.observe(labels, (performance.now() - start) / 1000);
      httpRequestsTotal.inc(labels);
    }
  };
}
