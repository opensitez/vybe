//! Shared .NET BCL frontend for all .NET-shaped compilers (VB, C#, F#, …).
//!
//! The .NET Base Class Library exposes the same namespace hierarchy regardless
//! of language: `System.Threading.Thread.Sleep`, `System.Diagnostics.Stopwatch`,
//! etc. This module is a single source of truth so every .NET compiler resolves
//! these identically.
//!
//! ## Module structure
//!
//! - **`resolver`** — the dotted-name resolution algorithm. Owns
//!   `DottedResolution`, `ResolutionContext`, `resolve_dotted_name`,
//!   `resolve_interface_call`, and the private import-prefix matching helpers.
//!
//! - **`imports`** — the implicit import list. `default_interface_imports()`
//!   returns the namespaces every .NET compiler auto-recognises (`System`,
//!   `System.Threading`, `System.Windows.Forms`, …). Language compilers
//!   `.extend()` it with their own additions.
//!
//! - **`namespaces`** — namespace-root recognition. `is_namespace_root` /
//!   `namespace_roots` answer "is `Math` a variable or the start of a
//!   namespace chain?".
//!
//! - **`host_map`** — `.NET → Vybe host` translation tables.
//!   `namespace_to_host_module` maps `system.console` → `wasi:cli`, and
//!   `map_host_func` maps `(wasi:cli, writeline)` → `log`.
//!
//! - **`types`** — type-related lookups & predicates: `known_types()`
//!   constructor table, `is_noop_method`, `is_known_constant`, and the
//!   PascalCase name-shape helpers.
//!
//! ## Future extensions
//!
//! When adding a real `Form` base class (and `Control` / `Button` / `TextBox`
//! as a class hierarchy that user code can `Inherits` from), add a new
//! `forms.rs` (and friends) sibling to these files. They use
//! `compiler_common::gui` helpers under the hood and get registered in the
//! type registry at host startup.
//!
//! Future framework frontends (MAUI, Flutter, Tkinter) follow the same
//! pattern: a sibling top-level module to `dotnet/`, structured the same way,
//! all delegating to `compiler_common::gui` for the canonical GUI emit.

pub mod resolver;
pub mod imports;
pub mod namespaces;
pub mod host_map;
pub mod types;
pub mod classes;

// ─── Public re-exports ───────────────────────────────────────────────────────
//
// The pre-split single-file `dotnet.rs` exposed everything at
// `compiler_common::dotnet::*`. The split keeps that flat surface for callers
// (compilers, walkers, host) — they continue to write
// `common::dotnet::resolve_dotted_name`, `common::dotnet::known_types`, etc.
// without caring which submodule the item lives in.

pub use resolver::{
    DottedResolution,
    ResolutionContext,
    resolve_dotted_name,
    resolve_interface_call,
};

pub use imports::default_interface_imports;

pub use namespaces::{
    is_namespace_root,
    namespace_roots,
};

pub use host_map::{
    namespace_to_host_module,
    map_host_func,
};

pub use types::{
    is_noop_method,
    is_known_constant,
    known_types,
    capitalize_control_name,
    capitalize_data_type,
};
