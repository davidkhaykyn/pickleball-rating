// Vercel serverless function: proxies the xlsx file from Google Drive
// avoiding CORS issues when fetching directly from the browser.

const FILE_ID = '1uYcmIttkWfH-rZ_KWdpuKoqnUHrk7cQn';

module.exports = async (req, res) => {
  try {
    const url = `https://drive.google.com/uc?export=download&id=${FILE_ID}`;
    let response = await fetch(url, { redirect: 'follow' });

    // For files that trigger Google's virus-scan confirmation, follow the
    // confirm token. (Not expected for a 200KB xlsx, but kept defensively.)
    const contentType = response.headers.get('content-type') || '';
    if (contentType.includes('text/html')) {
      const text = await response.text();
      const tokenMatch = text.match(/confirm=([0-9A-Za-z_-]+)/);
      if (tokenMatch) {
        const confirmUrl = `https://drive.google.com/uc?export=download&confirm=${tokenMatch[1]}&id=${FILE_ID}`;
        response = await fetch(confirmUrl, { redirect: 'follow' });
      } else {
        throw new Error('Drive returned HTML, no confirm token found');
      }
    }

    if (!response.ok) {
      throw new Error(`Drive responded ${response.status}`);
    }

    const buf = Buffer.from(await response.arrayBuffer());
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    // 60 seconds cache so rapid reloads don't hammer Drive,
    // but updates show within a minute of being entered.
    res.setHeader('Cache-Control', 'public, max-age=60, s-maxage=60');
    res.status(200).send(buf);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};

