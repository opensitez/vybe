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
use vybe_bytecode::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef,
};

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
fn memory_class(name: &'static str) -> ClassType {
    // `view_class` already installs the argc-branching `dotnet.span_ctor`
    // constructor; `new Memory<int>(array)` is the 1-arg (identity) branch.
    view_class(name).with_method(MethodDef::new(
        "Span",
        0,
        MethodBody::HostCall(HostTarget::new("ecma:object", "valueOf")),
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
    exports
}
