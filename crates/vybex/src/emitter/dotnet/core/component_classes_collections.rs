use super::super::super::class_exports::DotnetClassExport;
use super::component_classes_common::{common_constructor_class, constructor_class};
use vybe_bytecode::component_model::{ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        // .NET `List<T>` is shape-identical to ECMA-262 §23.1 Array.
        // The constructor materializes a real `ObjectKind::Array` via
        // `collections.new` (Op::ARRAY_NEW) and every method routes
        // through the corresponding `collections.*` primitive (which
        // itself routes to `ecma:array.*` per the WASM spec). The
        // .NET-name -> ECMA-name translation is the wrapper's job.
        collection_class_common(
            "dotnet.System.Collections.Generic",
            "List",
            "collections.new",
            &[
                ("Add", 1, "collections.push"),
                ("Remove", 1, "collections.remove"),
                ("RemoveAt", 1, "collections.remove_at"),
                ("Contains", 1, "collections.contains"),
                ("Count", 0, "collections.length"),
                ("Clear", 0, "collections.clear"),
                ("IndexOf", 1, "collections.index_of"),
                ("Sort", 0, "collections.sort"),
                ("Reverse", 0, "collections.reverse"),
                ("ToArray", 0, "collections.clone"),
                ("Item", 1, "collections.get"),
                ("Insert", 2, "collections.insert"),
                ("AddRange", 1, "collections.concat"),
            ],
        ),
        // .NET `Dictionary<K,V>` is shape-identical to ECMA-262 §24.1
        // `Map`. The wrapper materializes a real `ObjectKind::Map` via
        // `ecma:map/new` and forwards every method to the corresponding
        // `Map.prototype.*` host fn. No `vybe:types` involvement.
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("Dictionary")
                .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")))
                .with_method(MethodDef::new("Add", 2, MethodBody::HostCall(HostTarget::new("ecma:map", "set"))))
                .with_method(MethodDef::new("Item", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "get"))))
                .with_method(MethodDef::new("ContainsKey", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "has"))))
                .with_method(MethodDef::new("ContainsValue", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "containsValue"))))
                .with_method(MethodDef::new("Remove", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "delete"))))
                .with_method(MethodDef::new("Keys", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "keys"))))
                .with_method(MethodDef::new("Values", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "values"))))
                .with_method(MethodDef::new("Clear", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "clear"))))
                .with_method(MethodDef::new("Count", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "size")))),
        ),
        // .NET `Queue<T>` is a JS Array used FIFO - `Enqueue` appends
        // (push), `Dequeue` removes from the front (shift), `Peek`
        // looks at the front (`ecma:array.first`).
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("Queue")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new("Enqueue", 1, MethodBody::Common("collections.push".into())))
                .with_method(MethodDef::new("Dequeue", 0, MethodBody::Common("collections.shift".into())))
                .with_method(MethodDef::new("Peek", 0, MethodBody::HostCall(HostTarget::new("ecma:array", "first"))))
                .with_method(MethodDef::new("Count", 0, MethodBody::Common("collections.length".into())))
                .with_method(MethodDef::new("Clear", 0, MethodBody::Common("collections.clear".into())))
                .with_method(MethodDef::new("Contains", 1, MethodBody::Common("collections.contains".into())))
                .with_method(MethodDef::new("ToArray", 0, MethodBody::Common("collections.clone".into()))),
        ),
        // .NET `Stack<T>` is a JS Array used LIFO - `Push` appends
        // (push), `Pop` removes from the end (pop), `Peek` looks at
        // the end (`ecma:array.last`).
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("Stack")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new("Push", 1, MethodBody::Common("collections.push".into())))
                .with_method(MethodDef::new("Pop", 0, MethodBody::Common("collections.pop".into())))
                .with_method(MethodDef::new("Peek", 0, MethodBody::HostCall(HostTarget::new("ecma:array", "last"))))
                .with_method(MethodDef::new("Count", 0, MethodBody::Common("collections.length".into())))
                .with_method(MethodDef::new("Clear", 0, MethodBody::Common("collections.clear".into())))
                .with_method(MethodDef::new("Contains", 1, MethodBody::Common("collections.contains".into())))
                .with_method(MethodDef::new("ToArray", 0, MethodBody::Common("collections.clone".into()))),
        ),
        // .NET `HashSet<T>` is a real ECMA-262 §24.2 `Set`. Constructor
        // creates an `ObjectKind::Set`; methods route through the
        // matching `ecma:set.*` host fns.
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("HashSet")
                .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:set", "new")))
                .with_method(MethodDef::new("Add", 1, MethodBody::Common("dotnet.hashset_add".into())))
                .with_method(MethodDef::new("Remove", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "delete"))))
                .with_method(MethodDef::new("Contains", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "has"))))
                .with_method(MethodDef::new("Count", 0, MethodBody::HostCall(HostTarget::new("ecma:set", "size"))))
                .with_method(MethodDef::new("Clear", 0, MethodBody::HostCall(HostTarget::new("ecma:set", "clear"))))
                .with_method(MethodDef::new("UnionWith", 1, MethodBody::Common("dotnet.hashset_union_with".into())))
                .with_method(MethodDef::new("IntersectWith", 1, MethodBody::Common("dotnet.hashset_intersect_with".into())))
                .with_method(MethodDef::new("ExceptWith", 1, MethodBody::Common("dotnet.hashset_except_with".into())))
                .with_method(MethodDef::new("SymmetricExceptWith", 1, MethodBody::Common("dotnet.hashset_symmetric_except_with".into())))
                .with_method(MethodDef::new("IsSubsetOf", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "isSubsetOf"))))
                .with_method(MethodDef::new("IsSupersetOf", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "isSupersetOf"))))
                .with_method(MethodDef::new("Overlaps", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "overlaps")))),
        ),
        // `ConcurrentDictionary` is a thread-safe `Dictionary` - same
        // shape (ECMA Map). Atomicity isn't modeled; methods route the
        // same way.
        DotnetClassExport::new(
            "dotnet.System.Collections.Concurrent",
            ClassType::new("ConcurrentDictionary")
                .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")))
                .with_method(MethodDef::new("TryAdd", 2, MethodBody::HostCall(HostTarget::new("ecma:map", "set"))))
                .with_method(MethodDef::new("TryGetValue", 2, MethodBody::HostCall(HostTarget::new("ecma:map", "get"))))
                .with_method(MethodDef::new("AddOrUpdate", 3, MethodBody::HostCall(HostTarget::new("ecma:map", "set"))))
                .with_method(MethodDef::new("GetOrAdd", 2, MethodBody::HostCall(HostTarget::new("ecma:map", "get"))))
                .with_method(MethodDef::new("ContainsKey", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "has"))))
                .with_method(MethodDef::new("Remove", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "delete"))))
                .with_method(MethodDef::new("Clear", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "clear"))))
                .with_method(MethodDef::new("Count", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "size")))),
        ),
        // ConcurrentQueue / ConcurrentStack - same shape as their
        // non-concurrent counterparts (Array). Atomicity isn't
        // modeled at this layer.
        DotnetClassExport::new(
            "dotnet.System.Collections.Concurrent",
            ClassType::new("ConcurrentQueue")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new("Enqueue", 1, MethodBody::Common("collections.push".into())))
                .with_method(MethodDef::new("TryDequeue", 1, MethodBody::Common("collections.shift".into())))
                .with_method(MethodDef::new("TryPeek", 1, MethodBody::Common("collections.get".into())))
                .with_method(MethodDef::new("Count", 0, MethodBody::Common("collections.length".into()))),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Concurrent",
            ClassType::new("ConcurrentStack")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new("Push", 1, MethodBody::Common("collections.push".into())))
                .with_method(MethodDef::new("TryPop", 0, MethodBody::Common("collections.pop".into())))
                .with_method(MethodDef::new("TryPeek", 0, MethodBody::Common("collections.get".into())))
                .with_method(MethodDef::new("Count", 0, MethodBody::Common("collections.length".into()))),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("SortedDictionary")
                .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")))
                .with_method(MethodDef::new("Add", 2, MethodBody::HostCall(HostTarget::new("ecma:map", "set"))))
                .with_method(MethodDef::new("Item", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "get"))))
                .with_method(MethodDef::new("ContainsKey", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "has"))))
                .with_method(MethodDef::new("Remove", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "delete"))))
                .with_method(MethodDef::new("Keys", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "keys"))))
                .with_method(MethodDef::new("Values", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "values"))))
                .with_method(MethodDef::new("Clear", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "clear"))))
                .with_method(MethodDef::new("Count", 0, MethodBody::HostCall(HostTarget::new("ecma:map", "size")))),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("SortedSet")
                .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:set", "new")))
                .with_method(MethodDef::new("Add", 1, MethodBody::Common("dotnet.hashset_add".into())))
                .with_method(MethodDef::new("Remove", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "delete"))))
                .with_method(MethodDef::new("Contains", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "has"))))
                .with_method(MethodDef::new("Count", 0, MethodBody::HostCall(HostTarget::new("ecma:set", "size"))))
                .with_method(MethodDef::new("Clear", 0, MethodBody::HostCall(HostTarget::new("ecma:set", "clear")))),
        ),
        constructor_class("dotnet.System.Collections.Generic", "SortedList", "ecma:map", "new"),
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("LinkedList")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new("AddFirst", 1, MethodBody::Common("dotnet.linked_list_add_first".into())))
                .with_method(MethodDef::new("AddLast", 1, MethodBody::Common("dotnet.linked_list_add_last".into())))
                .with_method(MethodDef::new("Find", 1, MethodBody::Common("dotnet.linked_list_find".into())))
                .with_method(MethodDef::new("Clear", 0, MethodBody::Common("collections.clear".into())))
                .with_method(MethodDef::new("Count", 0, MethodBody::Common("collections.length".into()))),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections",
            ClassType::new("ArrayList")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new("Add", 1, MethodBody::Common("collections.push".into())))
                .with_method(MethodDef::new("Remove", 1, MethodBody::Common("collections.remove".into())))
                .with_method(MethodDef::new("RemoveAt", 1, MethodBody::Common("collections.remove_at".into())))
                .with_method(MethodDef::new("Contains", 1, MethodBody::Common("collections.contains".into())))
                .with_method(MethodDef::new("Count", 0, MethodBody::Common("collections.length".into())))
                .with_method(MethodDef::new("Capacity", 0, MethodBody::Common("collections.length".into())))
                .with_method(MethodDef::new("Clear", 0, MethodBody::Common("collections.clear".into())))
                .with_method(MethodDef::new("IndexOf", 1, MethodBody::Common("collections.index_of".into())))
                .with_method(MethodDef::new("IndexOf2", 2, MethodBody::Common("collections.index_of_from".into())))
                .with_method(MethodDef::new("LastIndexOf", 1, MethodBody::Common("collections.last_index_of".into())))
                .with_method(MethodDef::new("LastIndexOf2", 2, MethodBody::Common("collections.last_index_of_from".into())))
                .with_method(MethodDef::new("Sort", 0, MethodBody::Common("collections.sort".into())))
                .with_method(MethodDef::new("Reverse", 0, MethodBody::Common("collections.reverse".into())))
                .with_method(MethodDef::new("ReverseRange", 2, MethodBody::Common("collections.reverse_range".into())))
                .with_method(MethodDef::new("ToArray", 0, MethodBody::Common("collections.clone".into())))
                .with_method(MethodDef::new("Clone", 0, MethodBody::Common("collections.clone".into())))
                .with_method(MethodDef::new("Item", 1, MethodBody::Common("collections.get".into())))
                .with_method(MethodDef::new("Insert", 2, MethodBody::Common("collections.insert".into())))
                .with_method(MethodDef::new("InsertRange", 2, MethodBody::Common("collections.insert_range".into())))
                .with_method(MethodDef::new("RemoveRange", 2, MethodBody::Common("collections.remove_range".into())))
                .with_method(MethodDef::new("GetRange", 2, MethodBody::Common("collections.get_range".into())))
                .with_method(MethodDef::new("SetRange", 2, MethodBody::Common("collections.set_range".into())))
                .with_method(MethodDef::new("BinarySearch", 1, MethodBody::Common("collections.binary_search".into())))
                .with_method(MethodDef::new("AddRange", 1, MethodBody::Common("collections.concat".into()))),
            ),
        constructor_class("dotnet.System.Collections", "Hashtable", "ecma:map", "new"),
        common_constructor_class("dotnet.System.Collections", "Collection", "collections.new"),
    ]
}

fn collection_class_common(
    interface: &'static str,
    name: &'static str,
    ctor_common: &'static str,
    methods: &[(&'static str, u8, &'static str)],
) -> DotnetClassExport {
    let mut class = ClassType::new(name)
        .with_constructor(ConstructorDef::new(0).with_common_backing(ctor_common));
    for (method, arity, common) in methods {
        class = class.with_method(MethodDef::new(
            *method,
            *arity,
            MethodBody::Common((*common).into()),
        ));
    }
    DotnetClassExport::new(interface, class)
}

