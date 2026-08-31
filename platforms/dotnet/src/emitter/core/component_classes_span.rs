//! `System.Span<T>` / `System.ReadOnlySpan<T>` / `System.Memory<T>` /
//! `System.ReadOnlyMemory<T>` — the contiguous-view surface.
//!
//! .NET distinguishes these from arrays because they describe a window onto
//! memory the caller already owns. On this runtime there is no linear memory to
//! window: `stackalloc T[n]` and `new T[n]` both produce the same array, and a
//! Span IS that array.
//!
//! Declared once here, both C# and VB reach it through the common resolver.
//!
//! `Length` / indexing are deliberately absent: the receiver is already an
//! array, so they resolve without help.

use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef};

/// Members shared by every span/memory shape, as `(name, arity, module, func)`.
///
/// These are `HostCall`s rather than `collections.*` commons because the
/// call convention lines up exactly: a `HostCall` declared at the .NET arity
/// emits `module.func(receiver, ...args)` — the ECMA prototype-method shape.
/// The `collections.*` commons instead have FIXED stack contracts (`slice`
/// always wants `[array, start, end]`), so mapping a 1-arg .NET method onto
/// one yields malformed bytecode rather than a miss — `find_instance_method_on`
/// matches on exact arity and then emits blind.
///
/// Where the .NET signature and the ECMA one agree, that agreement is the
/// whole implementation:
///   `Slice(start)`  → `slice(arr, start)`  — both run to the end.
///   `ToArray()`     → `slice(arr)`         — both copy the whole thing.
///   `Fill(v)`       → `fill(arr, v)`       — both fill every slot.
const VIEW_HOST_METHODS: &[(&str, u8, &str, &str)] = &[
    ("Slice", 1, "ecma:array", "slice"),
    ("ToArray", 0, "ecma:array", "slice"),
    ("Fill", 1, "ecma:array", "fill"),
    ("Contains", 1, "ecma:array", "includes"),
    ("IndexOf", 1, "ecma:array", "indexOf"),
    ("LastIndexOf", 1, "ecma:array", "lastIndexOf"),
    ("Reverse", 0, "ecma:array", "reverse"),
];

/// Members that lower through a compiler-side emit rather than one host call.
///
/// `Slice(start, length)` is the one place .NET and ECMA disagree on meaning:
/// .NET's second operand is a LENGTH, ECMA's `slice` takes an END. That is
/// exactly `collections.get_range`'s `[array, index, count]` contract.
///
/// The `dotnet.span_*` entries need operands the call site never supplies
/// (`Clear` zeroes with `default(T)`), operands in the other order (`CopyTo`'s
/// receiver is the SOURCE), or a loop with no ECMA analogue. They compose the
/// shared primitives in `span_adapter.rs`; none of them adds a `collections.*`
/// name to the shared dispatcher.
const VIEW_COMMON_METHODS: &[(&str, u8, &str)] = &[
    ("Slice", 2, "dotnet.get_range_checked"),
    ("SequenceEqual", 1, "collections.sequence_equal"),
    ("IsEmpty", 0, "dotnet.span_is_empty"),
    ("Clear", 0, "dotnet.span_clear"),
    ("CopyTo", 1, "dotnet.span_copy_to"),
    ("TryCopyTo", 1, "dotnet.span_try_copy_to"),
    ("TrimStart", 1, "dotnet.span_trim_start"),
    ("TrimEnd", 1, "dotnet.span_trim_end"),
    ("Mismatch", 1, "dotnet.span_mismatch"),
];

fn view_class(name: &'static str) -> ClassType {
    let mut class = ClassType::new(name);
    for (method, arity, module, func) in VIEW_HOST_METHODS {
        class = class.with_method(MethodDef::new(
            *method,
            *arity,
            MethodBody::HostCall(HostTarget::new(*module, *func)),
        ));
    }
    for (method, arity, common) in VIEW_COMMON_METHODS {
        class = class.with_method(MethodDef::new(
            *method,
            *arity,
            MethodBody::Common((*common).into()),
        ));
    }
    // `new Span<int>(array)` windows an existing array — identity here, since a
    // span IS the array. `new Span<int>(array, start, length)` windows a
    // sub-range. A class carries ONE constructor slot and the emit passes the
    // real argument count, so both arities share a single argc-branching
    // adapter rather than two `ConstructorDef`s.
    class = class.with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.span_ctor"));
    class
}

/// `Memory<T>` / `ReadOnlyMemory<T>` add a `.Span` projection over the same
/// backing array. Both the constructor and `.Span` are IDENTITY: the
/// distinction .NET draws is about lifetime, which this runtime does not
/// model. `.Span` must NOT copy — `mem.Span[0]=77; mem.Span[0]` has to read
/// back 77, so every projection must hand back the same array.
///
/// `ecma:object.valueOf` IS that identity, per its spec default (§20.1.3.7
/// returns the object itself). It hands back the receiver's `Arc`, so the
/// array's identity — and every write through it — survives the projection.
/// ⛔ THE MEMORY SLICE IS DECLARED BEFORE THE VIEW'S. `tree_register` keeps the
/// FIRST declaration of an arity, so a `Slice` added after `view_class` would
/// never be reached — the shared `get_range_checked` would win and the ToString
/// ROLE would never be bound on the result.
fn memory_class(name: &'static str) -> ClassType {
    // `view_class` already installs the argc-branching `dotnet.span_ctor`
    // constructor; `new Memory<int>(array)` is the 1-arg (identity) branch.
    let mut class = ClassType::new(name)
        .with_method(MethodDef::new(
            "Slice",
            1,
            MethodBody::Common("dotnet.memory_slice".into()),
        ))
        .with_method(MethodDef::new(
            "Slice",
            2,
            MethodBody::Common("dotnet.memory_slice".into()),
        ));
    for (method, arity, module, func) in VIEW_HOST_METHODS {
        class = class.with_method(MethodDef::new(
            *method,
            *arity,
            MethodBody::HostCall(HostTarget::new(*module, *func)),
        ));
    }
    for (method, arity, common) in VIEW_COMMON_METHODS {
        class = class.with_method(MethodDef::new(
            *method,
            *arity,
            MethodBody::Common((*common).into()),
        ));
    }
    class
        .with_method(MethodDef::new(
            "Span",
            0,
            MethodBody::HostCall(HostTarget::new("ecma:object", "valueOf")),
        ))
        .with_method(MethodDef::new(
            "ToString",
            0,
            MethodBody::Common("dotnet.memory_to_string".into()),
        ))
}

/// `System.MemoryExtensions` — where .NET actually declares `AsSpan`,
/// `BinarySearch`, `Mismatch`, and friends, as extension methods over arrays,
/// strings, and spans. C# lets every one of them be written in either position,
/// so each needs BOTH a static entry here and an instance entry on the
/// receiver's own surface.
fn memory_extensions_class() -> ClassType {
    ClassType::new("MemoryExtensions")
        .with_method(MethodDef::static_method(
            "BinarySearch",
            2,
            MethodBody::Common("collections.binary_search".into()),
        ))
        .with_method(MethodDef::static_method(
            "Slice",
            2,
            MethodBody::HostCall(HostTarget::new("ecma:array", "slice")),
        ))
        // ⛔ THE WALKER SENDS EVERY `Slice` HERE. `m.Slice(0, 5)` is rewritten
        // to `MemoryExtensions.Slice(m, 0, 5)`, so the `ClassType` on
        // `ReadOnlyMemory` is never consulted — this is the only site that can
        // bind the ToString ROLE on the result, which is what makes it answer
        // through `var` rather than only through an explicit annotation.
        .with_method(MethodDef::static_method(
            "Slice",
            3,
            MethodBody::Common("dotnet.memory_slice".into()),
        ))
        .with_method(MethodDef::static_method(
            "ToArray",
            1,
            MethodBody::HostCall(HostTarget::new("ecma:array", "slice")),
        ))
        .with_method(MethodDef::static_method(
            "Fill",
            2,
            MethodBody::HostCall(HostTarget::new("ecma:array", "fill")),
        ))
        .with_method(MethodDef::static_method(
            "Contains",
            2,
            MethodBody::HostCall(HostTarget::new("ecma:array", "includes")),
        ))
        .with_method(MethodDef::static_method(
            "IndexOf",
            2,
            MethodBody::HostCall(HostTarget::new("ecma:array", "indexOf")),
        ))
        .with_method(MethodDef::static_method(
            "LastIndexOf",
            2,
            MethodBody::HostCall(HostTarget::new("ecma:array", "lastIndexOf")),
        ))
        .with_method(MethodDef::static_method(
            "Reverse",
            1,
            MethodBody::HostCall(HostTarget::new("ecma:array", "reverse")),
        ))
        .with_method(MethodDef::static_method(
            "SequenceEqual",
            2,
            MethodBody::Common("collections.sequence_equal".into()),
        ))
        .with_method(MethodDef::static_method(
            "IsEmpty",
            1,
            MethodBody::Common("dotnet.span_is_empty".into()),
        ))
        .with_method(MethodDef::static_method(
            "Clear",
            1,
            MethodBody::Common("dotnet.span_clear".into()),
        ))
        .with_method(MethodDef::static_method(
            "CopyTo",
            2,
            MethodBody::Common("dotnet.span_copy_to".into()),
        ))
        .with_method(MethodDef::static_method(
            "TryCopyTo",
            2,
            MethodBody::Common("dotnet.span_try_copy_to".into()),
        ))
        .with_method(MethodDef::static_method(
            "TrimStart",
            2,
            MethodBody::Common("dotnet.span_trim_start".into()),
        ))
        .with_method(MethodDef::static_method(
            "TrimEnd",
            2,
            MethodBody::Common("dotnet.span_trim_end".into()),
        ))
        .with_method(MethodDef::static_method(
            "AsSpan",
            1,
            MethodBody::HostCall(HostTarget::new("ecma:object", "valueOf")),
        ))
        .with_method(MethodDef::static_method(
            "AsSpan",
            3,
            MethodBody::Common("dotnet.get_range_checked".into()),
        ))
        .with_method(MethodDef::static_method(
            "Mismatch",
            2,
            MethodBody::Common("dotnet.span_mismatch".into()),
        ))
}

fn array_segment_class() -> ClassType {
    ClassType::new("ArraySegment")
        .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.array_segment_ctor"))
        .with_method(MethodDef::static_method(
            "Empty",
            0,
            MethodBody::Common("dotnet.array_segment_empty".into()),
        ))
        .with_method(MethodDef::static_method(
            "Get",
            2,
            MethodBody::Common("dotnet.array_segment_get".into()),
        ))
        .with_method(MethodDef::static_method(
            "Set",
            3,
            MethodBody::Common("dotnet.array_segment_set".into()),
        ))
        .with_method(MethodDef::static_method(
            "Slice",
            3,
            MethodBody::Common("dotnet.array_segment_slice".into()),
        ))
        .with_method(MethodDef::static_method(
            "CopyTo",
            2,
            MethodBody::Common("dotnet.array_segment_copy_to".into()),
        ))
        .with_method(MethodDef::static_method(
            "ToArray",
            1,
            MethodBody::Common("dotnet.array_segment_to_array".into()),
        ))
        .with_method(MethodDef::static_method(
            "Equals",
            2,
            MethodBody::Common("dotnet.array_segment_equals".into()),
        ))
        .with_method(MethodDef::new(
            "Item",
            1,
            MethodBody::Common("dotnet.array_segment_get".into()),
        ))
        .with_method(MethodDef::new(
            "SetItem",
            2,
            MethodBody::Common("dotnet.array_segment_set".into()),
        ))
        .with_method(MethodDef::new(
            "Slice",
            2,
            MethodBody::Common("dotnet.array_segment_slice".into()),
        ))
        .with_method(MethodDef::new(
            "CopyTo",
            1,
            MethodBody::Common("dotnet.array_segment_copy_to".into()),
        ))
        .with_method(MethodDef::new(
            "ToArray",
            0,
            MethodBody::Common("dotnet.array_segment_to_array".into()),
        ))
        .with_method(MethodDef::new(
            "Equals",
            1,
            MethodBody::Common("dotnet.array_segment_equals".into()),
        ))
}

fn array_pool_class() -> ClassType {
    ClassType::new("ArrayPool")
        .with_method(MethodDef::static_method(
            "Shared",
            0,
            MethodBody::Common("dotnet.array_pool_shared".into()),
        ))
        .with_method(MethodDef::new(
            "Rent",
            1,
            MethodBody::Common("dotnet.array_pool_rent".into()),
        ))
        .with_method(MethodDef::static_method(
            "Rent",
            1,
            MethodBody::Common("dotnet.array_pool_rent_static".into()),
        ))
        .with_method(MethodDef::new(
            "Return",
            1,
            MethodBody::Common("dotnet.array_pool_return".into()),
        ))
        .with_method(MethodDef::new(
            "Return",
            2,
            MethodBody::Common("dotnet.array_pool_return".into()),
        ))
}

/// `System.Buffers.Binary.BinaryPrimitives`.
///
/// ⛔ Registered as a TREE LEAF, not a synthesized class: the corpus writes the
/// fully-qualified `System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian`,
/// and only a leaf answers a dotted path.
pub fn binary_primitives_class() -> ClassType {
    ClassType::new("BinaryPrimitives")
                .with_method(MethodDef::static_method(
                    "ReadInt16LittleEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_i16_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadInt16BigEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_i16_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadUInt16LittleEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_u16_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadUInt16BigEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_u16_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadInt32LittleEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_i32_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadInt32BigEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_i32_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadUInt32LittleEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_u32_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadUInt32BigEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_u32_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadInt64LittleEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_i64_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadInt64BigEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_i64_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadUInt64LittleEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_u64_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadUInt64BigEndian",
                    1,
                    MethodBody::Common("dotnet.binprim_read_u64_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteInt16LittleEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_i16_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteInt16BigEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_i16_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteUInt16LittleEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_u16_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteUInt16BigEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_u16_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteInt32LittleEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_i32_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteInt32BigEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_i32_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteUInt32LittleEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_u32_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteUInt32BigEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_u32_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteInt64LittleEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_i64_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteInt64BigEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_i64_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteUInt64LittleEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_u64_le".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteUInt64BigEndian",
                    2,
                    MethodBody::Common("dotnet.binprim_write_u64_be".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ReverseEndianness",
                    1,
                    MethodBody::Common("dotnet.binprim_reverse_i32".into()),
                ))
}

fn memory_pool_class() -> ClassType {
    ClassType::new("MemoryPool")
        .with_method(MethodDef::static_method(
            "Shared",
            0,
            MethodBody::Common("dotnet.memory_pool_shared".into()),
        ))
        .with_method(MethodDef::new(
            "Rent",
            1,
            MethodBody::Common("dotnet.memory_pool_rent".into()),
        ))
        .with_method(MethodDef::static_method(
            "Rent",
            1,
            MethodBody::Common("dotnet.memory_pool_rent_static".into()),
        ))
}

fn memory_owner_class() -> ClassType {
    ClassType::new("MemoryPoolOwner").with_method(MethodDef::new(
        "Dispose",
        0,
        MethodBody::Common("dotnet.noop".into()),
    ))
}

/// `AsSpan` in instance position on a plain array (`data.AsSpan(1, 2)`).
///
/// Arrays never resolve against their own class — `lookup_instance_method`
/// sends every array/enumerable receiver to the shared `IEnumerable` surface —
/// so array extension methods have to be declared there. Kept to `AsSpan`: a
/// Span IS the array, so this is a no-op or a range, and nothing here changes
/// what `List<T>` and friends already resolve.
pub(super) fn add_array_extension_methods(class: &mut ClassType) {
    // Whole array → the same array, no copy (a span is a view, not a clone).
    class.methods.push(MethodDef::new(
        "AsSpan",
        0,
        MethodBody::HostCall(HostTarget::new("ecma:object", "valueOf")),
    ));
    // `AsSpan(start)` runs to the end — ECMA's `slice(start)` exactly.
    class.methods.push(MethodDef::new(
        "AsSpan",
        1,
        MethodBody::HostCall(HostTarget::new("ecma:array", "slice")),
    ));
    // `AsSpan(start, length)` — .NET's second operand is a LENGTH, so this is
    // `get_range`'s `[array, index, count]`, NOT ECMA's `slice(start, end)`.
    class.methods.push(MethodDef::new(
        "AsSpan",
        2,
        MethodBody::Common("dotnet.get_range_checked".into()),
    ));
}

pub(super) fn exports() -> Vec<DotnetClassExport> {
    let mut exports: Vec<DotnetClassExport> = ["Span", "ReadOnlySpan"]
        .into_iter()
        .map(|name| DotnetClassExport::new("dotnet.System", view_class(name)))
        .collect();
    exports.extend(
        ["Memory", "ReadOnlyMemory"]
            .into_iter()
            .map(|name| DotnetClassExport::new("dotnet.System", memory_class(name))),
    );
    exports.push(DotnetClassExport::new(
        "dotnet.System",
        memory_extensions_class(),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System",
        array_segment_class(),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Buffers",
        array_pool_class(),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Buffers.Binary",
        binary_primitives_class(),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Runtime.InteropServices",
        ClassType::new("MemoryMarshal").with_method(MethodDef::static_method(
            "CastBytes",
            3,
            MethodBody::Common("dotnet.memory_marshal_cast".into()),
        )),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Buffers",
        ClassType::new("ReadOnlySequence")
            .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.seq_new")),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Buffers",
        ClassType::new("SequenceReader")
            .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.seq_reader_new"))
            .with_method(MethodDef::new(
                "ReadNext",
                0,
                MethodBody::Common("dotnet.seq_read_next".into()),
            ))
            .with_method(MethodDef::new(
                "Remaining",
                0,
                MethodBody::Common("dotnet.seq_remaining".into()),
            ))
            .with_method(MethodDef::new(
                "Length",
                0,
                MethodBody::Common("dotnet.seq_remaining".into()),
            ))
            .with_method(MethodDef::new(
                "Consumed",
                0,
                MethodBody::Common("dotnet.seq_consumed".into()),
            )),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Buffers",
        ClassType::new("ArrayBufferWriter")
            .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.abw_new"))
            .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.abw_new"))
            .with_method(MethodDef::new(
                "GetSpan",
                1,
                MethodBody::Common("dotnet.abw_get_span".into()),
            ))
            .with_method(MethodDef::new(
                "GetMemory",
                1,
                MethodBody::Common("dotnet.abw_get_span".into()),
            ))
            .with_method(MethodDef::new(
                "Advance",
                1,
                MethodBody::Common("dotnet.abw_advance".into()),
            ))
            .with_method(MethodDef::new(
                "Clear",
                0,
                MethodBody::Common("dotnet.abw_clear".into()),
            ))
            .with_method(MethodDef::new(
                "WrittenCount",
                0,
                MethodBody::Common("dotnet.abw_written_count".into()),
            ))
            .with_method(MethodDef::new(
                "WrittenSpan",
                0,
                MethodBody::Common("dotnet.abw_written_span".into()),
            ))
            .with_method(MethodDef::new(
                "WrittenMemory",
                0,
                MethodBody::Common("dotnet.abw_written_span".into()),
            )),
    ));
    // `Marshal` — a tree leaf so the fully-qualified
    // `System.Runtime.InteropServices.Marshal.AllocHGlobal` resolves; the
    // synthesized interop classes beside it answer only the short name.
    exports.push(DotnetClassExport::new(
        "dotnet.System.Runtime.InteropServices",
        ClassType::new("Marshal")
            .with_method(MethodDef::static_method(
                "AllocHGlobal",
                1,
                MethodBody::Common("dotnet.marshal_alloc".into()),
            ))
            .with_method(MethodDef::static_method(
                "AllocCoTaskMem",
                1,
                MethodBody::Common("dotnet.marshal_alloc".into()),
            ))
            .with_method(MethodDef::static_method(
                "FreeHGlobal",
                1,
                MethodBody::Common("dotnet.marshal_free".into()),
            ))
            .with_method(MethodDef::static_method(
                "FreeCoTaskMem",
                1,
                MethodBody::Common("dotnet.marshal_free".into()),
            )),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Buffers",
        memory_pool_class(),
    ));
    exports.push(DotnetClassExport::new(
        "dotnet.System.Buffers",
        memory_owner_class(),
    ));
    exports
}
