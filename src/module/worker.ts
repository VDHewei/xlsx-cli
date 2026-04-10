import {readFileSync} from "fs";

import {assetsMap} from "./generated-assets";

const imageFavicon = readFileSync("./assets/favicon.ico", {encoding: "utf-8"});

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

type Asset = { mime: string; data: string };

const startServer = () => {
    // 构建本地内存服务器，处理静态资源请求
    const server = Bun.serve({
        port: 0, // 自动分配可用端口
        fetch(req) {
            const url = new URL(req.url);
            if (url.pathname.endsWith("favicon.ico")) {
                return new Response(imageFavicon, {headers: {"Content-Type": "image/x-icon"}});
            }
            if (assetsMap[url.pathname]) {
                const assets: Asset = assetsMap[url.pathname] as Asset;
                // 返回 base64 数据流给浏览器
                let base64Data: string;
                const values = assets.data.split(",");
                if (values.length >= 2 && values[1] !== undefined) {
                    base64Data = values[1];
                    const buffer = Buffer.from(base64Data, "base64");
                    return new Response(buffer, {headers: {"Content-Type": assets.mime}});
                }
            }
            return new Response(baseHTML, {headers: {"Content-Type": "text/html"}});
        },
    });
    console.log(`URL: ${server.url}`);
    self.postMessage({type: "PORT", port: server.port});
    return server;
}

if (!import.meta.main && process.env[`#start`] === undefined) {
    startServer();
    console.log(import.meta.main);
    process.env[`#start`] = `true`
}


