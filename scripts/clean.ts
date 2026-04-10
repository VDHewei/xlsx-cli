import {rmdirSync} from "node:fs";

const clean = () => {
    rmdirSync("./bin");
}

clean();