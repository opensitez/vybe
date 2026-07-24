//! Flutter platform package.
//!
//! A Flutter-shaped adapter over the existing `vybe_widgets`/`vybe:gui`
//! runtime. To Dart source we *are* Flutter (`Scaffold`, `Column`,
//! `Checkbox`, `Text`, …); under the hood every widget instantiates and
//! drives the same `vybe_widgets` controls that the dotnet (WinForms) and
//! plib (VCL) adapters already use — no Flutter-specific host functions,
//! no parallel widget runtime.
//!
//! The compiler-side code generation surface lives under [`emitter`].
//!
//! The adapter also owns its Dart *runtime* — the `runApp`/widget-tree
//! realizer that walks the constructed widget config objects and drives
//! `vybe:gui`. It is provided as source ([`runtime_source`]) and compiled into
//! a program ONLY when that program renders (references `runApp`), so
//! widget-only code (construction, `is`-checks, the TDD suite) carries none of
//! it — mirroring how the dotnet adapter emits per-class ctor chunks only for
//! the classes a program uses.

pub mod emitter;

/// The Flutter adapter's Dart runtime: `runApp`, the widget-tree realizer, and
/// the minimal `EdgeInsets`/`Alignment` value types. Pure Dart over the
/// existing `vybe:gui` host — no Flutter-specific host functions. The Dart
/// frontend appends this only when a module references `runApp`.
pub fn runtime_source() -> &'static str {
    include_str!("runtime.dart")
}
