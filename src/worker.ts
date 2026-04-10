import {readFileSync} from "fs";

const imageFavicon = readFileSync("./assets/favicon.ico", "utf-8");

const baseHTML = `
<html>
  <head>
    <link rel="shortcut icon" href="assets/favicon.ico" type="image/x-icon">
  </head>
    <body>
        <h1>Hello from bun v${Bun.version} !</h1>
    </body>
</html>
`;


// 构建本地内存服务器，处理静态资源请求
const server = Bun.serve({
    port: 0, // 自动分配可用端口
    fetch(req) {
        const url = new URL(req.url);
        if (url.pathname.endsWith("favicon.ico")) {
            return new Response(imageFavicon, {headers: {"Content-Type": "image/x-icon"}});
        }
        return new Response(baseHTML, {headers: {"Content-Type": "text/html"}});
        //if (url.pathname === "/bundle.js") {
        //    return new Response(bundleJS, { headers: { "Content-Type": "application/javascript" } });
        //}
        //if (url.pathname === "/bundle.css") {
        //    return new Response(bundleCSS, { headers: { "Content-Type": "text/css" } });
        //}
        //return new Response(indexHTML, { headers: { "Content-Type": "text/html" } });
    },
});
console.log(`URL: ${server.url}`);
self.postMessage({ type: "PORT", port: server.port });