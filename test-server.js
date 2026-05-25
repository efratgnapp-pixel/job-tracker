const http = require('http');
const PORT = process.env.PORT || 3001;
http.createServer((req, res) => {
  res.writeHead(200);
  res.end('hello from railway');
}).listen(PORT, () => {
  console.log('listening on port ' + PORT);
});
