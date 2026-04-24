// node:os — Node-style os module for Vybe.
//
// Re-exports system info primitives from `wasi:cli`.

export { platform, arch, cwd, machineName, userName, newLine } from "wasi:cli";
