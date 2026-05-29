//! Compilation target — detects whether we're on the Vybe VM or a standard WASM runtime.
//!
//! When targeting Vybe, compilers can use optimized host functions (vybe:array/range,
//! vybe:types/dictKeys, etc.) that are faster than pure-WASM equivalents.
//!
//! When targeting standard WASM, compilers emit portable bytecode sequences that work
//! on any compliant runtime (wasmtime, V8, wasm-micro-runtime).
//!
//! Usage:
//!   let target = Target::vybe();       // Use all Vybe extensions
//!   let target = Target::wasm();       // Pure WASM only
//!   let target = Target::detect(vm);   // Auto-detect from VM capabilities

/// Compilation target — controls what the compiler is allowed to emit.
#[derive(Debug, Clone)]
pub struct Target {
    /// Name of the target runtime.
    pub name: String,
    /// Whether Vybe-specific host functions are available.
    pub has_vybe_host: bool,
    /// Whether WASI imports are available (filesystem, CLI, etc).
    pub has_wasi: bool,
    /// Available host module prefixes (e.g., "ecma:math", "vybe:types").
    pub available_modules: Vec<String>,
}

impl Target {
    /// Vybe VM target — all host functions available.
    pub fn vybe() -> Self {
        Target {
            name: "vybe".into(),
            has_vybe_host: true,
            has_wasi: true,
            available_modules: vec![
                "ecma:math".into(), "ecma:string".into(), "ecma:number".into(),
                "ecma:array".into(), "ecma:object".into(), "ecma:json".into(),
                "ecma:map".into(), "ecma:set".into(),
                "ecma:weakmap".into(), "ecma:weakset".into(),
                "ecma:arraybuffer".into(), "ecma:sharedarraybuffer".into(),
                "ecma:dataview".into(), "ecma:fixedarray".into(),
                "ecma:structured-clone".into(), "ecma:value".into(), "ecma:date".into(),
                "ecma:int8array".into(), "ecma:uint8array".into(), "ecma:uint8clamped".into(),
                "ecma:int16array".into(), "ecma:uint16array".into(),
                "ecma:int32array".into(), "ecma:uint32array".into(),
                "ecma:float32array".into(), "ecma:float64array".into(),
                "ecma:bigint64array".into(), "ecma:biguint64array".into(),
                // ECMA-402 Intl
                "ecma:intl".into(),
                "ecma:intl/collator".into(),
                "ecma:intl/numberformat".into(),
                "ecma:intl/datetimeformat".into(),
                "ecma:intl/listformat".into(),
                "ecma:intl/pluralrules".into(),
                "ecma:intl/relativetimeformat".into(),
                "ecma:intl/segmenter".into(),
                "ecma:intl/locale".into(),
                "ecma:intl/displaynames".into(),
                "ecma:intl/durationformat".into(),
                // node:util — Node-aligned utility module
                "node:util".into(),
                "vybe:string".into(),
                "vybe:types".into(), "vybe:convert".into(), "vybe:object".into(),
                "ecma:regexp".into(), "vybe:crypto".into(),
                "wasi:sql/types".into(), "wasi:sql/readwrite".into(), "vybe:xml".into(),
                "vybe:drawing".into(),
                "wasi:cli".into(), "wasi:filesystem".into(), "wasi:http".into(),
                "wasi:random/random".into(), "wasi:random/insecure".into(), "wasi:random/insecure-seed".into(),
                "wasi:clocks".into(),
                "wasi:io/streams".into(), "wasi:io/poll".into(),
                "wasi:sockets/network".into(), "wasi:sockets/instance-network".into(),
                "wasi:sockets/tcp".into(), "wasi:sockets/tcp-create-socket".into(),
                "wasi:sockets/udp".into(), "wasi:sockets/udp-create-socket".into(),
                "wasi:sockets/ip-name-lookup".into(),
            ],
        }
    }

    /// Standard WASM target — no Vybe host functions, only WASM opcodes.
    pub fn wasm() -> Self {
        Target {
            name: "wasm".into(),
            has_vybe_host: false,
            has_wasi: false,
            available_modules: Vec::new(),
        }
    }

    /// WASI target — standard WASM + WASI imports, but no Vybe extensions.
    pub fn wasi() -> Self {
        Target {
            name: "wasi".into(),
            has_vybe_host: false,
            has_wasi: true,
            available_modules: vec![
                "wasi:cli".into(), "wasi:filesystem".into(), "wasi:http".into(),
                "wasi:random/random".into(), "wasi:random/insecure".into(), "wasi:random/insecure-seed".into(),
                "wasi:clocks".into(),
                "wasi:io/streams".into(), "wasi:io/poll".into(),
                "wasi:sockets/network".into(), "wasi:sockets/instance-network".into(),
                "wasi:sockets/tcp".into(), "wasi:sockets/tcp-create-socket".into(),
                "wasi:sockets/udp".into(), "wasi:sockets/udp-create-socket".into(),
                "wasi:sockets/ip-name-lookup".into(),
            ],
        }
    }

    /// Check if a specific host module is available.
    pub fn has_module(&self, module: &str) -> bool {
        self.available_modules.iter().any(|m| m == module)
    }
}

/// Default target is Vybe (full host support).
impl Default for Target {
    fn default() -> Self {
        Self::vybe()
    }
}
