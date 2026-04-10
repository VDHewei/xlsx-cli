import { Webview } from "webview-bun";

const worker = new Worker(new URL("./worker.ts", import.meta.url));

// 使用 Promise 等待 Worker 传递端口号
const waitForPort = (): Promise<number> => {
    return new Promise((resolve) => {
        worker.addEventListener("message", (event) => {
            if (event.data.type === "PORT") {
                resolve(event.data.port);
            }
        });
    });
};

const start = async (): Promise<void> => {
    // 1. 阻塞等待 Worker 启动服务器并返回可用端口
    // 必须在 run() 之前完成，否则 run() 会卡死事件循环导致 onmessage 无法触发
    const port = await waitForPort();
    const indexUrl = `http://localhost:${port}/index.html`;
    console.log("Server ready, navigating to:", indexUrl);
    // 2. 初始化 Webview
    const webview = new Webview();
    webview.title = "webview-bun";
    // 3. 使用原生 navigate 方法导航，不要用 eval
    webview.navigate(indexUrl);
    // 4. 监听 worker 退出
    worker.addEventListener("exit", () => {
        process.exit(0);
    });
    // 5. 启动 UI 阻塞循环
    // 运行到此行后，主线程将被 webview 接管，直到窗口关闭
    webview.run();
    // 6. run() 结束（窗口关闭），清理 worker
    setTimeout(() => {
        worker.terminate();
    }, 100);
};

await start();