//! Capability-based host access control (WASI security model).
//!
//! A [`Capabilities`] set is the sandbox policy for a VM: which host-function
//! groups the runtime is permitted to register/expose (fs, sockets, http, …).
//! The host wiring consults it at registration time. Lives in `vybe_runtime`
//! because the capability model is a VM-level primitive, independent of any
//! particular host-function crate.

use std::collections::HashSet;

/// Capability flags for host module access.
/// Follows WASI's capability-based security model.
#[derive(Debug, Clone)]
pub struct Capabilities {
    granted: HashSet<Capability>,
}

/// Individual capabilities that can be granted or denied.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Capability {
    /// Console I/O (stdout, stderr). Safe for most contexts.
    Console,
    /// Runtime compilation from source text (guest-visible eval / compile).
    /// Static project loading and compile-time-resolved imports do not use it.
    DynamicCompile,
    /// Filesystem read access.
    FileRead,
    /// Filesystem write access (implies FileRead).
    FileWrite,
    /// Network: outbound HTTP requests.
    Http,
    /// Network: TCP/UDP sockets (server + client).
    Sockets,
    /// Database connections (SQLite, MySQL, etc.).
    Database,
    /// Environment variables and process args.
    Environment,
    /// GUI / window creation.
    Gui,
    /// Threading / background tasks.
    Threading,
    /// Cryptographic operations.
    Crypto,
    /// System clock access (time, sleep).
    Clock,
    /// Random number generation.
    Random,
    /// XML parsing.
    Xml,
    /// HTTP server (binding ports, handling requests). Required for `vybex --serve`
    /// and any script calling `vybe:http/server.listen`.
    HttpServer,
    /// Spawning child processes (`node:child_process.{spawnSync, execSync,
    /// execFileSync}`, `node:process.kill`). Carries OS-level escape
    /// potential — gated separately from FileWrite.
    Process,
}

impl Capabilities {
    /// Full access — all capabilities granted. For trusted CLI usage.
    pub fn all() -> Self {
        use Capability::*;
        let mut granted = HashSet::new();
        for cap in [
            Console,
            DynamicCompile,
            FileRead,
            FileWrite,
            Http,
            Sockets,
            Database,
            Environment,
            Gui,
            Threading,
            Crypto,
            Clock,
            Random,
            Xml,
            HttpServer,
            Process,
        ] {
            granted.insert(cap);
        }
        Capabilities { granted }
    }

    /// Safe subset — only pure computation, no I/O or side effects.
    /// Suitable for untrusted code (web playground, sandboxed eval).
    pub fn safe() -> Self {
        use Capability::*;
        let mut granted = HashSet::new();
        for cap in [Console, Clock, Random] {
            granted.insert(cap);
        }
        Capabilities { granted }
    }

    /// No capabilities — pure computation only.
    pub fn none() -> Self {
        Capabilities {
            granted: HashSet::new(),
        }
    }

    /// Custom: start with none, add specific capabilities.
    pub fn with(caps: &[Capability]) -> Self {
        Capabilities {
            granted: caps.iter().copied().collect(),
        }
    }

    pub fn has(&self, cap: Capability) -> bool {
        self.granted.contains(&cap)
    }

    pub fn grant(&mut self, cap: Capability) {
        self.granted.insert(cap);
    }

    pub fn revoke(&mut self, cap: Capability) {
        self.granted.remove(&cap);
    }
}
