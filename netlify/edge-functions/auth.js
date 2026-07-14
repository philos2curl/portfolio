// Netlify Edge Function — runs on the server before the page is delivered.
// Checks for a valid auth cookie. If missing or wrong, redirects to /password.html.
// Password is stored as a Netlify environment variable (PORTFOLIO_PASSWORD).

export default async function(request, context) {
  const url = new URL(request.url);
  const password = Deno.env.get('PORTFOLIO_PASSWORD') || 'JerelPM';

  // Parse cookies from request header
  const cookieHeader = request.headers.get('Cookie') || '';
  const cookies = Object.fromEntries(
    cookieHeader.split(';')
      .map(c => c.trim())
      .filter(Boolean)
      .map(c => {
        const [k, ...v] = c.split('=');
        return [k.trim(), decodeURIComponent(v.join('=').trim())];
      })
  );

  // Valid cookie — serve the page normally
  if (cookies['portfolio_auth'] === password) {
    return context.next();
  }

  // No valid cookie — redirect to password page, preserving destination
  const next = encodeURIComponent(url.pathname);
  return Response.redirect(
    new URL(`/password.html?next=${next}`, url.origin),
    302
  );
}
