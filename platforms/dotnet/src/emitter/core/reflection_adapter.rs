//! `System.Reflection` surfaces that are DATA rather than machinery.
//!
//! The reflection machinery is `vybe_compiler::primitives::reflection` — a
//! compile-time `ReflectionBinding` resolver that already answers `typeof(T)`,
//! `.Name`, `.FullName`, `GetMethod`, `GetField`, `GetParameters` and
//! `MakeGenericMethod`. It reads its type metadata from the component
//! descriptors this crate exports, so what a type needs in order to be
//! reflectable is a REGISTRATION here, not new code there.
//!
//! These three are the ones the resolver cannot derive from a descriptor:
//! `OpCodes` is a table of constants, and `CustomAttributeData` /
//! `NullabilityInfoContext` are ordinary objects.

use vybe_compiler::primitives::class_slots::{self, ObjSource, ValueSource};
use vybe_compiler::primitives::collections;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use super::object_fields::field_slot;

/// The CIL opcodes the corpus names, as `(member, name, value)`.
///
/// `OpCode.Name` is the CIL mnemonic — dotted, lowercase (`ldarg.0`), NOT the
/// C# member spelling (`Ldarg_0`) — and `Value` is the one-byte encoding from
/// ECMA-335 Partition VI.
const OPCODES: &[(&str, &str, i32)] = &[
    ("Nop", "nop", 0x00),
    ("Ldarg_0", "ldarg.0", 0x02),
    ("Ldarg_1", "ldarg.1", 0x03),
    ("Ldarg_2", "ldarg.2", 0x04),
    ("Ldarg_3", "ldarg.3", 0x05),
    ("Ldloc_0", "ldloc.0", 0x06),
    ("Ldloc_1", "ldloc.1", 0x07),
    ("Ldloc_2", "ldloc.2", 0x08),
    ("Ldloc_3", "ldloc.3", 0x09),
    ("Stloc_0", "stloc.0", 0x0A),
    ("Stloc_1", "stloc.1", 0x0B),
    ("Stloc_2", "stloc.2", 0x0C),
    ("Stloc_3", "stloc.3", 0x0D),
    ("Ldnull", "ldnull", 0x14),
    ("Ldc_I4_0", "ldc.i4.0", 0x16),
    ("Ldc_I4_1", "ldc.i4.1", 0x17),
    ("Dup", "dup", 0x25),
    ("Pop", "pop", 0x26),
    ("Ret", "ret", 0x2A),
    ("Add", "add", 0x58),
];

/// Every `(member, name, value)` this module can answer.
pub fn opcodes() -> &'static [(&'static str, &'static str, i32)] {
    OPCODES
}

/// The `dotnet.` common that reads `OpCodes.<member>`.
pub fn opcode_common(member: &str) -> String {
    format!("{OPCODE_COMMON_PREFIX}{member}")
}

pub const OPCODE_COMMON_PREFIX: &str = "dotnet.opcode.";

/// `OpCodes.<member>` — an `OpCode` carrying its mnemonic and its encoding.
///
/// Stack: `[]` → `[opcode]`.
pub fn emit_opcode(chunk: &mut Chunk, member: &str, line: u32) {
    let Some((_, name, value)) = OPCODES
        .iter()
        .find(|(spelling, _, _)| spelling.eq_ignore_ascii_case(member))
    else {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    };
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_string_const("OpCode", line);
    set_field(chunk, "__type", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_string_const(name, line);
    set_field(chunk, "Name", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_f64_const(*value as f64, line);
    set_field(chunk, "Value", line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
}

/// `CustomAttributeData.GetCustomAttributes(target)` — the attributes declared
/// on `target`, as a list.
///
/// The descriptors carry no attribute data for platform types, so a platform
/// type answers the empty list — which is what .NET answers for a type with no
/// custom attributes, and is a list rather than null either way.
///
/// Stack: `[target]` → `[array]`.
pub fn emit_custom_attribute_data(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
}

/// `New NullabilityInfoContext()` — the per-call-site cache .NET uses to read
/// nullability annotations. Nothing is cached here; the object exists so the
/// members hang off something.
///
/// Stack: `[args…]` → `[context]`.
pub fn emit_nullability_info_context_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_string_const("NullabilityInfoContext", line);
    set_field(chunk, "__type", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
}

fn set_field(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}
