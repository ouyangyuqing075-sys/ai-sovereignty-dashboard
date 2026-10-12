import { cp, mkdir, rm } from "node:fs/promises";

const output = new URL("../docs/", import.meta.url);
const root = new URL("../", import.meta.url);

await rm(output, { recursive: true, force: true });
await mkdir(output, { recursive: true });
await cp(new URL("index.html", root), new URL("index.html", output));
await cp(new URL("src", root), new URL("src", output), { recursive: true });
await cp(new URL("data", root), new URL("data", output), { recursive: true });

console.log("Built GitHub Pages files in docs/");
