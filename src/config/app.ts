
type AppConfig = {
    title: string;
    icon: string;
    version: string;
    defaultLang: string;
    supportedLangs: string[];
};

const App: AppConfig = {
    title: "xlsx-cli",
    icon: '$favicon.icon',
    version: "1.0.0",
    defaultLang: "zh-CN",
    supportedLangs: ["zh-CN", "zh-TW", "en", "ja"],
};

export {
    App,
};
