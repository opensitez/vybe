//! PHP adapter — bytecode-only.
//!
//! Mirrors the `emitter/dotnet/` pattern: PHP-specific class surfaces
//! (`DateTime`, `DateTimeImmutable`, `DateInterval`, ...) are
//! described as `emit_*` opcode-emitter functions that compose
//! pre-existing host fns (`ecma:date.*`, `ecma:string.*`, etc.) into
//! the PHP-shaped surface.
//!
//! No new host fns are registered. Everything is bytecode emission
//! over the spec-conformant ECMA-262 / WASI surface.
//!
//! The PHP profile binds language fns to `common:php.<name>` emit
//! targets; `emitter::dispatch::emit_common` routes those names to
//! the `emit_*` functions in this module.

pub mod component_classes;
pub mod copy_adapter;
pub mod datetime_adapter;

// Inline-emit adapters for PHP top-level functions. Each `emit_<fn>`
// takes `(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32)`
// and emits WASM opcodes directly into `chunks[current]` — same
// shape as `emitter/dotnet/core/datetime_adapter.rs` and
// `emitter/php/datetime_adapter.rs`. Routed through
// `emitter::dispatch::emit_common` via `common:php.<name>` and
// bound from the PHP profile's `[builtins]` section. No JS
// polyfills; no PHP-specific host fns.
pub mod array_adapter;
pub mod autoload_adapter;
pub mod call_adapter;
pub mod ctype_adapter;
pub mod db_adapter;
pub mod dispatch;
pub mod empty_adapter;
pub mod env_adapter;
pub mod error_adapter;
pub mod fiber_adapter;
pub mod filesystem_adapter;
pub mod http_adapter;
pub mod ini_adapter;
pub mod json_adapter;
pub mod math_adapter;
pub mod mysqli_adapter;
pub mod numeric_adapter;
pub mod output_adapter;
pub mod packing_adapter;
pub mod pdo_adapter;
pub mod reflection_adapter;
pub mod relational_adapter;
pub mod runtime_adapter;
pub mod session_adapter;
pub mod serialization_adapter;
pub mod spl_adapter;
pub mod string_adapter;
pub mod superglobal_adapter;
pub mod type_guard;
pub mod url_adapter;
pub mod version_adapter;
pub mod weak_adapter;
pub mod xml_adapter;
