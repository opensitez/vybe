//! `java.lang.Enum` runtime metadata — the class-name → constants map.
//!
//! `EnumSet.allOf(Color.class)` has to reach `Color`'s declaration-ordered
//! constant names, and all it is handed is a STRING: `X.class` lowers to the
//! type's name, an invariant shared code depends on (`builtins.rs`
//! `symbol_require` — "Java: `X.class` is a string, so `Class.forName` must
//! agree with it", and .NET represents a type the same way). A tree leaf
//! receives already-compiled stack values, so it cannot turn that string back
//! into the class global at compile time the way `symbol_require` does.
//!
//! So the enum publishes its own metadata. Each enum's static initializer
//! calls `java.lang.Enum.__vybe_declare(name, names)`, which stores the list
//! under the class name in one module global. That is not a workaround for the
//! string — it is what `Class.getEnumConstants()` actually is, and it is the
//! only shape that also answers a name computed at runtime
//! (`Class.forName("Color")`).
//!
//! The alternative — having the walker prepend a compile-time name array at
//! every `EnumSet.*` call site — is what this replaces: it made a `java.util`
//! class the property of one frontend, so Kotlin (and any other consumer of
//! `jvm.java.*`) reached nothing.

use vybe_compiler::primitives::{globals, instructions::host};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Module global holding `{ "Color": ["RED","GREEN"], … }`.
pub const REGISTRY_GLOBAL: &str = "__vybe_jvm_enum_constants";

/// Push the registry, creating and storing an empty object on first touch.
/// Stack on exit: `[registry]`.
fn emit_registry(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    globals::emit_read(&mut chunks[current], REGISTRY_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let undef = chunks[current].add_import("wasm:js-undefined", "test");
    chunks[current].emit_call(undef, 1, line);
    chunks[current].emit_if(line);
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    globals::emit_write(&mut chunks[current], REGISTRY_GLOBAL, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

/// `java.lang.Enum.__vybe_declare(name, names)` — stack `[name, names]` → `[names]`.
///
/// Called from the enum's own `__static_init_block__`, ahead of the constant
/// constructions, so the list is published before any code can ask for it.
pub fn emit_declare(chunks: &mut [Chunk], current: usize, line: u32) {
    let names = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, names, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);

    emit_registry(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, names, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, names, line);
}

/// `Class.getEnumConstants()` — stack `[class]` → `[names]`.
///
/// A class the program never declared yields `undefined`, which is what the
/// JDK's `null` return for a non-enum class means here.
pub fn emit_constants_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let name = chunks[current].alloc_scratch(1);
    emit_class_name(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
    emit_registry(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

/// Reduce whatever a JVM frontend calls a `java.lang.Class` to its NAME.
///
/// The two frontends do not agree, and neither is wrong for its own language:
/// Java's `X.class` is the name string (the invariant shared `symbol_require`
/// is built on), while Kotlin's `X::class.java` is an object carrying `name`
/// /`canonicalName`/`simpleName`, because Kotlin code reads those properties.
/// Reconciling them here — at the JVM boundary, where `java.lang.Class` lives
/// — is what lets one leaf serve both, instead of either walker changing its
/// own surface to suit the other.
fn emit_class_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_string_const("name", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_end(line);
}
