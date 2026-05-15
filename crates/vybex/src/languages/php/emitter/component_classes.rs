//! PHP `ClassType` exports — placeholder for descriptor-driven
//! resolution analogous to `emitter/dotnet/core/component_classes.rs`.
//!
//! For now the PHP class surface is reached via
//! `[known_types]` / `[builtins]` profile entries that route to the
//! `common:php.*` emit names registered in
//! `emitter::dispatch::emit_common`. Future descriptor-based dispatch
//! (mirroring `dotnet::lookup_constructor` /
//! `dotnet::lookup_instance_method`) can collect entries here.
