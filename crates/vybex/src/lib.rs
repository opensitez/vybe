pub mod cli;
pub mod gui_launch;
pub mod server;

// The eval / dynamic-compile layer lives in `vybe_compiler` (the only crate
// below the language crates that can call the compiler). Re-exported so the
// shell's call sites keep writing `crate::dynamic` / `crate::adapters` / etc.
pub use vybe_compiler::{adapters, dynamic, host_imports};
