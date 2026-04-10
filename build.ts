import {rcedit} from "rcedit";
import {mkdirSync, writeFileSync, copyFileSync, readFileSync, existsSync, renameSync} from "node:fs";
import {join, resolve} from "node:path";
import dotenv from "dotenv";

dotenv.config({quiet: true});
// ==========================================
// 1. 构建配置项
// ==========================================
const APP_NAME = "webview-cli";
const OUT_DIR = "./bin";
const ICON_DIR = "./assets"; // 需提前准备好 icon.ico (Win), icon.icns (Mac), icon.png (Linux)
const ICON_FILE = "xlsx-1";
const DOMAIN = "vdhewei.io";
const ENTRY_FILES = ["./src/index.ts", /*"./src/worker.ts"*/];
const VERSION: string = JSON.parse(readFileSync("package.json", {encoding: 'utf-8'})).version;
// ==========================================

// 2. 核心构建逻辑
// ==========================================
function build() {
    const platform = process.platform; // 'win32' | 'darwin' | 'linux'
    // const args = process.argv;
    // console.log("args:", args);
    const outPath = join(OUT_DIR, platform);
    mkdirSync(outPath, {recursive: true});
    console.log(`[1/3] 🚀 开始为 ${platform} 平台编译 Bun 二进制文件...`);
    const iconPath = resolve(join(ICON_DIR, `${ICON_FILE}.ico`));
    // Step 1: 调用 Bun 编译裸二进制文件
    const binName = platform === "win32" ? `${APP_NAME}.exe` : APP_NAME;
    const binPath = join(outPath, binName);
    const compileCmd = ["bun", "build",
        "--compile",
        "--minify",
    ]
    const flag = process.env[`WINDOWS_HIDE_CONSOLE`];
    if (flag !== undefined && (flag as string).toLowerCase() === "true") {
        compileCmd.push("--windows-hide-console");
    }
    const argList = [
        "--windows-icon", iconPath,
        "--windows-title", APP_NAME,
        "--windows-version", VERSION,
        "--windows-copyright", DOMAIN,
        "--compile-autoload-tsconfig",
        "--compile-autoload-package-json",
        ...ENTRY_FILES, "--outfile", binPath
    ];
    compileCmd.push(...argList);
    console.log(compileCmd.join(' '));
    const compileResult = Bun.spawnSync(compileCmd);
    if (compileResult.exitCode !== 0) {
        console.error("编译失败:", compileResult.stderr.toString());
        process.exit(1);
    }
    console.log(`[1/3] ✅ 编译成功: ${binPath}`);
    // Step 2: 根据平台注入 Title 和 Icon
    console.log("[2/3] 🎨 注入图标和应用标题...");
    if (platform === "win32") {
        injectWinMetadata(binPath, compileCmd);
    } else if (platform === "darwin") {
        injectMacMetadata(binPath, outPath);
    } else if (platform === "linux") {
        injectLinuxMetadata(binPath, outPath);
    }
    console.log("[3/3] 🎉 构建与后处理全部完成！");
}

// ==========================================
// 3. 平台特定后处理实现
function checkDisableConsole(arg: string[]): boolean {
    for (const flag of arg) {
        if (flag.trim() === "--windows-hide-console") {
            return true;
        }
    }
    console.log("windows-console-enable");
    return false;
}

// ==========================================
/** Windows: 使用 rcedit 注入图标与元数据 */
function injectWinMetadata(binPath: string, ...args: any[]) {
    const iconPath = resolve(join(ICON_DIR, `${ICON_FILE}.ico`));
    if (!existsSync(iconPath)) return console.warn("未找到 icon.ico，跳过图标注入");
    // 调用 rcedit (需系统已安装 rcedit 并在 PATH 中)
    //  const result = Bun.spawnSync([
    //      "rcedit", binPath,
    //      "--set-icon", iconPath,
    //      "--set-version-string", `FileDescription=${APP_NAME}`,
    //      "--set-version-string", `ProductName=${APP_NAME}`
    //  ]);
    rcedit(binPath, {
        icon: iconPath,
        'version-string': {
            FileDescription: APP_NAME,
            ProductName: APP_NAME
        }
    }).then(() => {
        console.info("rcedit 执行成功 ✅ ");
        if (args[0] instanceof Array &&
            (args[0] as string[]).length > 0 && checkDisableConsole(args[0] as string[])) {
            consoleDisableForWindows(binPath);
        }
    }).catch((reason) => {
        if (reason !== undefined) {
            console.error("rcedit 执行失败 ", reason);
        } else {
            console.error("rcedit 执行失败，可能需要安装或以管理员权限运行");
        }
    });
}

/** macOS: 构建 .app 标准目录结构 */
function injectMacMetadata(binPath: string, outPath: string) {
    const appBundlePath = join(outPath, `${APP_NAME}.app`);
    const contentsDir = join(appBundlePath, "Contents");
    const macOSDir = join(contentsDir, "MacOS");
    const resourcesDir = join(contentsDir, "Resources");
    const domain = [...DOMAIN].reverse().join('');
    mkdirSync(macOSDir, {recursive: true});
    mkdirSync(resourcesDir, {recursive: true});
    // 移动二进制文件到 MacOS 目录
    const newBinPath = join(macOSDir, APP_NAME);
    renameSync(binPath, newBinPath); // 注意：这里用 rename 相当于移动文件
    // 复制图标
    const iconPath = resolve(join(ICON_DIR, `${ICON_FILE}.icns`));
    if (existsSync(iconPath)) {
        copyFileSync(iconPath, join(resourcesDir, "icon.icns"));
    }
    // 生成 Info.plist (定义 Title 和 Icon)
    const plistContent = `<?xml version="1.0" encoding="UTF-8"?>
	<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
	<plist version="1.0">
	<dict>
	    <key>CFBundleExecutable</key> <string>${APP_NAME}</string>
	    <key>CFBundleName</key> <string>${APP_NAME}</string>
	    <key>CFBundleIconFile</key> <string>icon</string>
	    <key>CFBundleIdentifier</key> <string>${domain}.${APP_NAME.toLowerCase()}</string>
	    <key>CFBundleVersion</key> <string>1.0.0</string>
	</dict>
	</plist>`;
    writeFileSync(join(contentsDir, "Info.plist"), plistContent);
}

/** Linux: 生成 .desktop 文件 (图标需放置在标准图标目录，这里输出到同目录作为便携版) */
function injectLinuxMetadata(binPath: string, outPath: string) {
    const iconPath = resolve(join(ICON_DIR, `${ICON_FILE}.png`));
    const destIcon = join(outPath, `${APP_NAME}.png`);
    if (existsSync(iconPath)) copyFileSync(iconPath, destIcon);
    const desktopContent = `[Desktop Entry]
	Type=Application
	Name=${APP_NAME}
	Exec=${binPath}
	Icon=${destIcon}
	Categories=Utility;`;
    writeFileSync(join(outPath, `${APP_NAME}.desktop`), desktopContent);
}

function consoleDisableForWindows(exePath: string) {
    try {
        const buf = readFileSync(exePath);
        // 1. 验证 DOS 头 Magic (MZ)
        if (buf.toString('ascii', 0, 2) !== 'MZ') {
            throw new Error("不是一个有效的 PE (EXE) 文件");
        }
        // 2. 获取 PE 头偏移量 (e_lfanew)
        const peOffset = buf.readUInt32LE(0x3c);
        // 3. 验证 PE 签名 (PE\0\0)
        if (buf.toString('ascii', peOffset, peOffset + 4) !== 'PE\0\0') {
            throw new Error("PE 签名无效");
        }
        // 4. Optional Header 起始偏移量 = PE签名(4) + COFF头(20)
        const optionalHeaderOffset = peOffset + 24;
        const magic = buf.readUInt16LE(optionalHeaderOffset);
        // 5. 关键修正：无论是 PE32 还是 PE32+，Subsystem 在 Optional Header 中的偏移量都是 68 (0x44)
        const subsystemOffset = optionalHeaderOffset + 68;
        const subsystem = buf.readUInt16LE(subsystemOffset);
        console.log(`文件解析: 架构=${magic === 0x020b ? 'PE32+ (64位)' : magic === 0x010b ? 'PE32 (32位)' : '未知'}, 当前 Subsystem 值=${subsystem}`);
        if (subsystem === 3) { // 3 = IMAGE_SUBSYSTEM_WINDOWS_CUI (控制台程序)
            buf.writeUInt16LE(2, subsystemOffset); // 2 = IMAGE_SUBSYSTEM_WINDOWS_GUI (GUI程序)
            writeFileSync(exePath, buf);
            console.log("✅ 成功: 已将 Subsystem 从 CONSOLE(3) 修改为 GUI(2)，cmd 黑窗口已被彻底隐藏。");
        } else if (subsystem === 2) {
            console.log("ℹ️ 提示: 当前 Subsystem 已经是 GUI(2)，无需修改。");
        } else {
            console.log("⚠️ 警告: 未知的 Subsystem 值:", subsystem);
        }
    } catch (e) {
        console.error("❌ 修改 exe 失败:", (e as Error).message);
    }
}

// 执行构建
build();