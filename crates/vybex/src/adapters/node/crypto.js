// node:crypto — Node-style crypto module for Vybe.
//
// Re-exports hash primitives. Node's `crypto.createHash('sha256')
// .update(data).digest('hex')` shape requires local functions in the
// adapter; today we expose the flat hash functions directly.

export { sha256, md5 } from "vybe:crypto";
