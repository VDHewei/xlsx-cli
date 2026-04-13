import { Webview } from "webview-bun";
import { App } from "@config/app.ts";
import { runCli } from "@module/cli.ts";

// 判断是否 CLI 模式：如果带任何 option/command 参数
const args = process.argv.slice(2);
const isCliMode = args.length > 0;

if (isCliMode) {
    // CLI 模式：执行命令后退出
    runCli(process.argv);
} else {
    // UI 模式：启动 webview
    const worker = new Worker(Bun.resolveSync("@module/worker", import.meta.dir));

    const waitForPort = (): Promise<number> => {
        return new Promise((resolve, reject) => {
            worker.addEventListener("message", (event) => {
                if (event.data.type === "PORT") {
                    resolve(event.data.port);
                }
            });
            worker.addEventListener("error", (err) => {
                console.error("Worker error:", err);
                reject(err);
            });
        });
    };

    const start = async (): Promise<void> => {
        const webview = new Webview();
        const port = await waitForPort();
        const indexUrl = `http://localhost:${port}/index.html`;
        console.log("Server ready, navigating to:", indexUrl);

        webview.title = App.title;
        webview.navigate(indexUrl);

        worker.addEventListener("exit", () => {
            process.exit(0);
        });

        webview.run();

        setTimeout(() => {
            worker.terminate();
        }, 100);
    };

    await start();
}
