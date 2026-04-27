//! # `web:*` host imports — WHATWG / W3C web platform APIs.
//!
//! Web platform APIs that complement ECMA-262 but live outside the
//! language spec. Every JavaScript runtime ships these; Vybe exposes
//! them under the `web:*` namespace so language profiles can target a
//! standard surface instead of inventing `vybe:*` adapters.
//!
//! - `web:crypto`         — Web Cryptography API (`crypto.randomUUID`,
//!                          `crypto.getRandomValues`, `crypto.subtle.digest`)
//! - `web:url`            — WHATWG URL (`URL`, `URLSearchParams`)
//! - `web:encoding`       — WHATWG Encoding (`TextEncoder`, `TextDecoder`)
//! - `web:fetch`          — WHATWG Fetch (`fetch`, `Request`, `Response`,
//!                          `Headers`)
//! - `web:timers`         — HTML Timers (`setTimeout`, `clearTimeout`,
//!                          `setInterval`, `clearInterval`)

pub mod crypto;
pub mod url;
pub mod encoding;
pub mod fetch;
pub mod timers;

use vybe_bytecode::VM;

pub fn register(vm: &mut VM) {
    crypto::register(vm);
    url::register(vm);
    encoding::register(vm);
    fetch::register(vm);
    timers::register(vm);
}
