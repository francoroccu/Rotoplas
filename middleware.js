import { next } from '@vercel/functions';

// Protege todo el sitio con usuario y contraseña (HTTP Basic Auth).
// Configurar DASHBOARD_USER y DASHBOARD_PASSWORD en Vercel › Settings › Environment Variables.
export default function middleware(request) {
  const user = process.env.DASHBOARD_USER;
  const pass = process.env.DASHBOARD_PASSWORD;
  if (!user || !pass) {
    return new Response('Falta configurar DASHBOARD_USER y DASHBOARD_PASSWORD en Vercel.', { status: 503 });
  }
  const auth = request.headers.get('authorization') || '';
  const [scheme, encoded] = auth.split(' ');
  if (scheme === 'Basic' && encoded) {
    const decoded = atob(encoded);
    const i = decoded.indexOf(':');
    if (i > -1 && decoded.slice(0, i) === user && decoded.slice(i + 1) === pass) {
      return next();
    }
  }
  return new Response('Acceso restringido', {
    status: 401,
    headers: { 'WWW-Authenticate': 'Basic realm="Rotoplas Peru", charset="UTF-8"' },
  });
}
