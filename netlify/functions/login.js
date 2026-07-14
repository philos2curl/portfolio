// Netlify Serverless Function — handles password form submission.
// Checks the submitted password server-side and sets an auth cookie on success.

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  const params = new URLSearchParams(event.body);
  const submitted = params.get('password') || '';
  const next = params.get('next') || '/';
  const password = process.env.PORTFOLIO_PASSWORD || 'JerelPM';

  if (submitted === password) {
    // Correct — set auth cookie valid for 24 hours, redirect to destination
    return {
      statusCode: 302,
      headers: {
        'Set-Cookie': `portfolio_auth=${encodeURIComponent(password)}; Path=/; SameSite=Strict; Max-Age=86400`,
        'Location': next,
        'Cache-Control': 'no-store',
      },
      body: '',
    };
  }

  // Wrong — redirect back to password page with error flag
  return {
    statusCode: 302,
    headers: {
      'Location': `/password.html?next=${encodeURIComponent(next)}&error=1`,
      'Cache-Control': 'no-store',
    },
    body: '',
  };
};
