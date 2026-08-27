//! `System.Collections.Specialized`, `…ObjectModel` wrappers, and
//! `PriorityQueue` — the collection types this platform had never declared.
//!
//! Every one of these was `undefined is not callable` from all three .NET
//! frontends, so none of it is a C#-only gap.

use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef,
};

fn inst(name: &'static str, argc: u8, common: &'static str) -> MethodDef {
    MethodDef::new(name, argc, MethodBody::Common(common.into()))
}

fn inst_host(name: &'static str, argc: u8, iface: &'static str, func: &'static str) -> MethodDef {
    MethodDef::new(name, argc, MethodBody::HostCall(HostTarget::new(iface, func)))
}

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        // ── BitVector32 ───────────────────────────────────────────────────
        //
        // A STRUCT wrapping a single `int`, so the value IS the int and `Data`
        // reads it back. Giving it a boxed representation would buy nothing:
        // every documented member is defined in terms of that one word.
        DotnetClassExport::new(
            "dotnet.System.Collections.Specialized",
            ClassType::new("BitVector32")
                .with_constructor(ConstructorDef::new(1).with_common_backing("collections.identity"))
                .with_method(inst("Data", 0, "collections.identity"))
                .with_method(MethodDef::static_method(
                    "CreateMask",
                    0,
                    MethodBody::Common("dotnet.bitvector32_create_mask".into()),
                ))
                // §CreateMask(previous) — the NEXT mask up. Declared alongside
                // the zero-arg form because .NET's documented idiom chains
                // them, and an arity-1 call would otherwise resolve to the
                // arity-0 leaf and answer 1 forever.
                .with_method(MethodDef::static_method(
                    "CreateMask",
                    1,
                    MethodBody::Common("dotnet.bitvector32_create_mask_next".into()),
                )),
        ),
        // ── NameValueCollection ───────────────────────────────────────────
        //
        // A missing key READS NULL here — unlike `Dictionary`, which throws —
        // so `Item` deliberately does NOT reuse `dotnet.dict_get_or_throw`.
        DotnetClassExport::new(
            "dotnet.System.Collections.Specialized",
            ClassType::new("NameValueCollection")
                .with_constructor(
                    ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")),
                )
                // ⛔ `Add` APPENDS here — see `emit_name_value_add`. Only
                // `Set` replaces.
                .with_method(inst("Add", 2, "dotnet.name_value_add"))
                .with_method(inst_host("Set", 2, "ecma:map", "set"))
                .with_method(inst_host("Item", 1, "ecma:map", "get"))
                .with_method(inst_host("Get", 1, "ecma:map", "get"))
                .with_method(inst_host("Remove", 1, "ecma:map", "delete"))
                .with_method(inst_host("Clear", 0, "ecma:map", "clear"))
                .with_method(inst_host("Count", 0, "ecma:map", "size"))
                .with_method(inst_host("AllKeys", 0, "ecma:map", "keys"))
                .with_method(inst_host("Keys", 0, "ecma:map", "keys")),
        ),
        // ── OrderedDictionary ─────────────────────────────────────────────
        //
        // Indexable by POSITION and by key, and BOTH go through native
        // indexing — the declared `Item` member never runs. See
        // `specialized_adapter::emit_ordered_dictionary_add` for why the
        // backing stores each value twice.
        //
        // `Remove`/`Keys`/`Values` are NOT declared: removing would have to
        // undo both halves of the storage, and a member that half-works is
        // worse than one that resolves to nothing.
        DotnetClassExport::new(
            "dotnet.System.Collections.Specialized",
            ClassType::new("OrderedDictionary")
                .with_constructor(
                    ConstructorDef::new(0).with_common_backing("dotnet.ordered_dictionary_new"),
                )
                .with_method(inst("Add", 2, "dotnet.ordered_dictionary_add"))
                .with_method(inst("Count", 0, "collections.length")),
        ),
        // ── ReadOnlyCollection / ReadOnlyDictionary ───────────────────────
        //
        // A VIEW, not a copy: §the wrapper reflects later changes to the
        // wrapped collection. Identity is therefore the correct construction,
        // and cloning would be the bug — it would silently freeze a snapshot.
        // Read-only-ness is a compile-time restriction in C#, so nothing is
        // lost by the wrapper being the collection itself.
        DotnetClassExport::new(
            "dotnet.System.Collections.ObjectModel",
            ClassType::new("ReadOnlyCollection")
                .with_constructor(ConstructorDef::new(1).with_common_backing("collections.identity"))
                .with_method(inst("Count", 0, "collections.length"))
                .with_method(inst("Item", 1, "dotnet.list_get_checked"))
                .with_method(inst("Contains", 1, "collections.contains"))
                .with_method(inst("IndexOf", 1, "collections.index_of"))
                .with_method(inst("ToArray", 0, "collections.clone")),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.ObjectModel",
            ClassType::new("ReadOnlyDictionary")
                .with_constructor(ConstructorDef::new(1).with_common_backing("collections.identity"))
                .with_method(inst_host("Count", 0, "ecma:map", "size"))
                .with_method(inst("Item", 1, "dotnet.dict_get_or_throw"))
                .with_method(inst_host("ContainsKey", 1, "ecma:map", "has"))
                .with_method(inst_host("Keys", 0, "ecma:map", "keys"))
                .with_method(inst_host("Values", 0, "ecma:map", "values")),
        ),
        // ── PriorityQueue<TElement, TPriority> ────────────────────────────
        //
        // Dequeues the SMALLEST priority, and .NET does not promise stability
        // among equal priorities, so a linear scan for the minimum is a
        // conforming implementation — the heap is an optimisation, not part of
        // the contract.
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("PriorityQueue")
                .with_constructor(
                    ConstructorDef::new(0).with_common_backing("dotnet.priority_queue_new"),
                )
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.priority_queue_new"),
                )
                .with_method(inst("Enqueue", 2, "dotnet.priority_queue_enqueue"))
                .with_method(inst("Dequeue", 0, "dotnet.priority_queue_dequeue"))
                .with_method(inst("Peek", 0, "dotnet.priority_queue_peek"))
                .with_method(inst("Count", 0, "dotnet.priority_queue_count"))
                .with_method(inst("Clear", 0, "dotnet.priority_queue_clear")),
        ),
    ]
}
