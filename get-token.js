const { google } = require('googleapis');
const http = require('http');
const url = require('url');

const CLIENT_ID     = process.env.GOOGLE_OAUTH_CLIENT_ID     || 'YOUR_CLIENT_ID';
const CLIENT_SECRET = process.env.GOOGLE_OAUTH_CLIENT_SECRET || 'YOUR_CLIENT_SECRET';
const REDIRECT_URI  = 'http://localhost:3000/oauth2callback';

const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

const authUrl = oauth2Client.generateAuthUrl({
  access_type: 'offline',
  scope: ['https://www.googleapis.com/auth/drive'],
  prompt: 'consent',
});

console.log('\n🔗 아래 URL을 브라우저에서 열고 snu.lacrosse@gmail.com 으로 로그인하세요:\n');
console.log(authUrl);
console.log('\n⏳ 인증 대기 중...\n');

// 콜백 서버
const server = http.createServer(async (req, res) => {
  const { pathname, query } = url.parse(req.url, true);
  if (pathname !== '/oauth2callback') return;

  const code = query.code;
  if (!code) {
    res.end('코드가 없어요.');
    return;
  }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    res.end('<h2>✅ 인증 완료! 터미널로 돌아가세요.</h2>');
    server.close();

    console.log('\n✅ 토큰 발급 성공!\n');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('Railway 환경변수에 아래 값들을 추가하세요:');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log(`GOOGLE_OAUTH_CLIENT_ID=${CLIENT_ID}`);
    console.log(`GOOGLE_OAUTH_CLIENT_SECRET=${CLIENT_SECRET}`);
    console.log(`GOOGLE_OAUTH_REFRESH_TOKEN=${tokens.refresh_token}`);
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
  } catch (e) {
    res.end(`오류: ${e.message}`);
    console.error('토큰 발급 실패:', e.message);
    server.close();
  }
});

server.listen(3000, () => {});
