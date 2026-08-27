//! `System.Collections.Immutable` — persistent collections.
//!
//! The whole surface is copy-on-write: every "mutation" returns a NEW
//! collection and leaves the receiver untouched. That is the ONE fact these
//! types add over their mutable counterparts, so it is expressed once, in the
//! adapter (`immutable_adapter`), rather than by giving each type its own
//! bespoke `Add`. `ImmutableArray.Add` and `ImmutableQueue.Enqueue` are the
//! same emit with different .NET names.
//!
//! Storage is deliberately the ordinary array/map/set representation, not a
//! private one: an immutable collection is a normal collection that nobody is
//! allowed to write through. Giving it a distinct backing would fork every
//! `Count`, indexer, and iteration path for no observable gain.

use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{ClassType, HostTarget, MethodBody, MethodDef};

const IFACE: &'static str = "dotnet.System.Collections.Immutable";

/// A static factory or constant — `.NET` spells these as properties
/// (`ImmutableList<T>.Empty`) and as generic factories
/// (`ImmutableArray.Create<T>(…)`). Both arrive here as static members; a
/// zero-arg static IS the property read (`argc == 0` is the read), which is
/// why `Empty` needs no separate node kind.
fn stat(name: &'static str, argc: u8, common: &'static str) -> MethodDef {
    MethodDef::static_method(name, argc, MethodBody::Common(common.into()))
}

fn inst(name: &'static str, argc: u8, common: &'static str) -> MethodDef {
    MethodDef::new(name, argc, MethodBody::Common(common.into()))
}

/// An instance member that reuses an existing host leaf verbatim.
///
/// The map and set members below are the SAME endpoints `Dictionary` and
/// `HashSet` already resolve to. An immutable collection is not a different
/// container, so `ContainsKey` must not acquire a second implementation that
/// can drift from the first — only the members that COPY are new.
fn inst_host(name: &'static str, argc: u8, iface: &'static str, func: &'static str) -> MethodDef {
    MethodDef::new(name, argc, MethodBody::HostCall(HostTarget::new(iface, func)))
}

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        // ── ImmutableArray / ImmutableList ────────────────────────────────
        //
        // Both are sequences over the same storage. They differ in .NET only
        // by the length member's NAME — `Length` on the array (it is a struct
        // wrapping T[]), `Count` on the list — so both are declared and both
        // answer from the same emit.
        DotnetClassExport::new(
            IFACE,
            ClassType::new("ImmutableArray")
                .with_method(stat("Create", 0, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 1, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 2, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 3, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 4, "dotnet.immutable_seq_create"))
                .with_method(stat("CreateRange", 1, "collections.clone"))
                .with_method(stat("Empty", 0, "dotnet.immutable_seq_empty"))
                .with_method(inst("Add", 1, "dotnet.immutable_seq_add"))
                .with_method(inst("AddRange", 1, "dotnet.immutable_seq_add_range"))
                .with_method(inst("RemoveAt", 1, "dotnet.immutable_seq_remove_at"))
                .with_method(inst("SetItem", 2, "dotnet.immutable_seq_set_item"))
                // ⛔ `Length` IS NOT DECLARED. A declared instance member is
                // reachable from receivers that are not of this type, and
                // `Length` is the one C# string member that collides — every
                // `someString.Length` in the language resolved to THIS leaf
                // and answered `undefined` (measured tree-wide; `Count` is
                // safe because a C# string has no `.Count`). The array backing
                // answers `.Length` natively anyway, so declaring it bought
                // nothing and cost the string surface.
                .with_method(inst("Count", 0, "collections.length"))
                .with_method(inst("IsEmpty", 0, "dotnet.immutable_is_empty"))
                .with_method(inst("Item", 1, "dotnet.list_get_checked"))
                .with_method(inst("ToArray", 0, "collections.clone")),
        ),
        DotnetClassExport::new(
            IFACE,
            ClassType::new("ImmutableList")
                .with_method(stat("Create", 0, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 1, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 2, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 3, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 4, "dotnet.immutable_seq_create"))
                .with_method(stat("CreateRange", 1, "collections.clone"))
                .with_method(stat("CreateBuilder", 0, "dotnet.immutable_seq_empty"))
                .with_method(stat("Empty", 0, "dotnet.immutable_seq_empty"))
                .with_method(inst("Add", 1, "dotnet.immutable_seq_add"))
                .with_method(inst("AddRange", 1, "dotnet.immutable_seq_add_range"))
                .with_method(inst("RemoveAt", 1, "dotnet.immutable_seq_remove_at"))
                .with_method(inst("SetItem", 2, "dotnet.immutable_seq_set_item"))
                .with_method(inst("Count", 0, "collections.length"))
                .with_method(inst("IsEmpty", 0, "dotnet.immutable_is_empty"))
                .with_method(inst("Item", 1, "dotnet.list_get_checked"))
                .with_method(inst("ToArray", 0, "collections.clone")),
        ),
        // ── The builder — the ONE mutable type here ───────────────────────
        //
        // ⛔ NOT named `Builder`. .NET spells it as the NESTED type
        // `ImmutableList<T>.Builder`, but this tree keys types by their BARE
        // name in a flat map, so registering `Builder` shadowed every user
        // class called `Builder` — measured: three fluent-builder tests broke
        // because `new Builder().Add(...)` resolved to this type instead of
        // theirs. This is the `Text`-is-a-Flutter-widget-and-a-.NET-name
        // collision `namespaceplan.md` describes, and the fix it prescribes is
        // a name that cannot be shadowed rather than a scoped lookup.
        //
        // `CreateBuilder()` hands back plain array storage, so the builder's
        // `Add` is the ordinary mutating push and `ToImmutable` is a COPY.
        // The copy is not ceremony: without it the built list would alias the
        // builder, and a later `builder.Add` would mutate a value .NET
        // guarantees frozen.
        DotnetClassExport::new(
            IFACE,
            ClassType::new("ImmutableListBuilder")
                .with_method(inst("Add", 1, "dotnet.list_add"))
                .with_method(inst("Count", 0, "collections.length"))
                .with_method(inst("Item", 1, "dotnet.list_get_checked"))
                .with_method(inst("ToImmutable", 0, "collections.clone")),
        ),
        // ── ImmutableDictionary ───────────────────────────────────────────
        DotnetClassExport::new(
            IFACE,
            ClassType::new("ImmutableDictionary")
                .with_method(stat("Empty", 0, "dotnet.immutable_map_empty"))
                .with_method(stat("Create", 0, "dotnet.immutable_map_empty"))
                .with_method(stat("CreateBuilder", 0, "dotnet.immutable_map_empty"))
                .with_method(inst("Add", 2, "dotnet.immutable_map_add"))
                .with_method(inst("SetItem", 2, "dotnet.immutable_map_add"))
                .with_method(inst_host("ContainsKey", 1, "ecma:map", "has"))
                .with_method(inst_host("Count", 0, "ecma:map", "size"))
                .with_method(inst("Item", 1, "dotnet.dict_get_or_throw")),
        ),
        // ── ImmutableHashSet ──────────────────────────────────────────────
        DotnetClassExport::new(
            IFACE,
            ClassType::new("ImmutableHashSet")
                .with_method(stat("Create", 0, "dotnet.immutable_set_create"))
                .with_method(stat("Create", 1, "dotnet.immutable_set_create"))
                .with_method(stat("Create", 2, "dotnet.immutable_set_create"))
                .with_method(stat("Create", 3, "dotnet.immutable_set_create"))
                .with_method(stat("Create", 4, "dotnet.immutable_set_create"))
                .with_method(stat("Empty", 0, "dotnet.immutable_set_empty"))
                .with_method(inst("Add", 1, "dotnet.immutable_set_add"))
                .with_method(inst("Remove", 1, "dotnet.immutable_set_remove"))
                .with_method(inst("Contains", 1, "sets.has"))
                .with_method(inst("Union", 1, "dotnet.immutable_set_union"))
                .with_method(inst("Intersect", 1, "dotnet.immutable_set_intersect"))
                .with_method(inst("Except", 1, "dotnet.immutable_set_except"))
                .with_method(inst("Count", 0, "sets.size"))
                .with_method(inst("IsEmpty", 0, "dotnet.immutable_set_is_empty")),
        ),
        // ── ImmutableQueue / ImmutableStack ───────────────────────────────
        //
        // Same storage, opposite ends. A queue peeks the FRONT and a stack
        // peeks the BACK, which is the only reason these are two entries and
        // not one — `Enqueue` and `Push` are the identical append.
        DotnetClassExport::new(
            IFACE,
            ClassType::new("ImmutableQueue")
                .with_method(stat("Empty", 0, "dotnet.immutable_seq_empty"))
                .with_method(stat("Create", 0, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 1, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 2, "dotnet.immutable_seq_create"))
                .with_method(inst("Enqueue", 1, "dotnet.immutable_seq_add"))
                .with_method(inst("Dequeue", 0, "dotnet.immutable_queue_dequeue"))
                .with_method(inst("Peek", 0, "dotnet.immutable_queue_peek"))
                .with_method(inst("IsEmpty", 0, "dotnet.immutable_is_empty"))
                .with_method(inst("Count", 0, "collections.length")),
        ),
        DotnetClassExport::new(
            IFACE,
            ClassType::new("ImmutableStack")
                .with_method(stat("Empty", 0, "dotnet.immutable_seq_empty"))
                .with_method(stat("Create", 0, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 1, "dotnet.immutable_seq_create"))
                .with_method(stat("Create", 2, "dotnet.immutable_seq_create"))
                .with_method(inst("Push", 1, "dotnet.immutable_seq_add"))
                .with_method(inst("Pop", 0, "dotnet.immutable_stack_pop"))
                .with_method(inst("Peek", 0, "dotnet.immutable_stack_peek"))
                .with_method(inst("IsEmpty", 0, "dotnet.immutable_is_empty"))
                .with_method(inst("Count", 0, "collections.length")),
        ),
    ]
}
