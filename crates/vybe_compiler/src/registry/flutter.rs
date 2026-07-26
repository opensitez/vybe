//! Flutter widget adapters — common-resolver integration.
//!
//! When a module imports `package:flutter/*`, mount `flutter` as an ambient
//! namespace-tree root so unqualified widget names (`Scaffold(...)`,
//! `Text('x')`) resolve to their `flutter.<name>` tree `Type` and construct
//! through the ONE common-resolver `Ctor` path
//! (`Compiler::emit_tree_ctor_construction`). The widget catalog itself is
//! DATA registered in `platforms/flutter` (`emitter::tree_register`); nothing
//! Flutter-specific is installed on the compiler — no per-module ctor globals,
//! no host functions. Mirrors how `.NET` surfaces are reached via the shared
//! tree, not a bespoke resolver.

use crate::ast::{ImportKind, Module};
use crate::compiler::Compiler;

/// True when the module imports any `package:flutter/*` library.
pub(crate) fn module_uses_flutter(module: &Module) -> bool {
    module.imports.iter().any(|import| match &import.kind {
        ImportKind::Simple { path, .. }
        | ImportKind::Named { path, .. }
        | ImportKind::Wildcard { path, .. }
        // `dart:ui` is the lower half of the same surface — `Rect`, `Color`,
        // `Offset`, `Size`, `Paint`, `Path` all live in the Flutter catalog, so
        // a module importing it needs the same ambient root. Widget code
        // reaches these types through `package:flutter/…`; `dart:ui` code
        // imports them directly.
        | ImportKind::Default { path, .. } => {
            path.starts_with("package:flutter/") || path == "dart:ui"
        }
    })
}

impl Compiler {
    /// Make the `flutter` tree root ambient for this module — unqualified
    /// widget names resolve under it via the common resolver.
    pub(crate) fn mount_flutter_ambient(&mut self) {
        self.mount_ambient_root("flutter");
    }
}
