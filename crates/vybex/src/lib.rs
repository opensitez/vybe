// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));

pub mod cli;
pub mod dap;
pub mod debug_repl;
pub mod gui_launch;
pub mod server;
pub mod watch;

// The eval / dynamic-compile layer lives in `vybe_compiler` (the only crate
// below the language crates that can call the compiler). Re-exported so the
// shell's call sites keep writing `crate::dynamic` / `crate::adapters` / etc.
pub use vybe_compiler::{adapters, dynamic, host_imports};
