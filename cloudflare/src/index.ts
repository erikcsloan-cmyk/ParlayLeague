export default {
  async fetch(request: Request): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return Response.json({
        status: "ok",
        application: "Parlay League",
        platform: "Cloudflare Workers"
      });
    }

    return new Response(
      `<!DOCTYPE html>
      <html lang="en">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>Parlay League</title>
        </head>
        <body style="font-family:Arial;text-align:center;padding:60px 20px">
          <h1>Parlay League</h1>
          <h2>Cloudflare Edition</h2>
          <p>Our new application is under construction.</p>
          <p>Cloudflare Worker is running successfully!</p>
        </body>
      </html>`,
      { headers: { "Content-Type": "text/html; charset=utf-8" } }
    );
  }
};