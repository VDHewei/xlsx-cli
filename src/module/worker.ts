import {assetsMap} from "./generated-assets";

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

type AssetResult = {
    buffer: Buffer,
    asset: Asset,
};

const getAssetsData = (name: string): AssetResult | undefined => {
    const assets = assetsMap[name] as Asset;
    if(assets === undefined || assets.data === undefined || assets.data === null){
        return undefined;
    }
    // 返回 base64 数据流给浏览器
    let base64Data: string;
    const values = assets.data.split(",");
    if (values.length >= 2 && values[1] !== undefined) {
        base64Data = values[1];
        return {
            asset: assets,
            buffer: Buffer.from(base64Data, "base64"),
        }
    }
    return undefined;
}

const startServer = () => {
    // 构建本地内存服务器，处理静态资源请求
    const server = Bun.serve({
        port: 0, // 自动分配可用端口
        fetch(req) {
            const url = new URL(req.url);
            if (url.pathname.endsWith("favicon.ico")) {
                const fav = getAssetsData(`favicon.ico`);
                if (fav !== undefined) {
                    console.log(url.pathname, fav.asset.mime);
                    return new Response(fav.buffer, {headers: {"Content-Type": "image/x-icon"}});
                }
            }
            if (assetsMap[url.pathname]) {
                const assets = getAssetsData(url.pathname);
                // 返回 base64 数据流给浏览器
                if (assets !== undefined) {
                    return new Response(assets.buffer, {headers: {"Content-Type": assets.asset.mime}});
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


