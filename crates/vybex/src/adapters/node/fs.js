// node:fs — Node-style filesystem module for Vybe.
//
// Re-exports `wasi:filesystem` primitives under Node's specifier.
// Full Node-API-shape translation (callbacks, promises, sync/async
// variants, Buffer handling) is follow-up work that needs local
// functions in the adapter.

export { readFile, writeFile, appendFile, exists, remove, rename, copy, mkdir, listDir, stat, isDir, isFile } from "wasi:filesystem";
