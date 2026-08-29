//! JS standard collection surface — the `wasm:js-*` imports Vybe uses
//! to expose JS-canonical collections (Array / Object / Map / Set /
//! WeakMap / WeakSet / ArrayBuffer / DataView / 11 typed-arrays).
//! Marshaling contract pinned in `JS_BUILTIN_CONVENTIONS.md`.

pub mod canon_builtins;
pub mod js_array_builtins;
pub mod js_arraybuffer_builtins;
pub mod js_fixedarray_builtins;
pub mod js_json_builtins;
pub mod js_map_builtins;
pub mod js_object_builtins;
pub mod js_primitive_builtins;
pub mod js_set_builtins;
pub mod js_string_builtins;
pub mod js_structured_clone;
pub mod js_typedarray_builtins;
pub mod js_weakmap_builtins;
