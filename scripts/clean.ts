import {rmSync} from "node:fs";

const clean = () => {
    // 清理 bin 目录 ，递归删除所有文件和子目录
    console.log('clean bin');
    rmSync('bin', {recursive: true});
}

clean();