use super::super::super::class_exports::DotnetClassExport;
use super::component_classes_common::constructor_class;
use vybe_runtime::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef,
};

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
            "dotnet.list_new",
            &[
                ("Add", 1, "dotnet.list_add"),
                ("Remove", 1, "collections.remove"),
                ("RemoveAll", 1, "dotnet.list_remove_all"),
                ("RemoveAt", 1, "collections.remove_at"),
                ("Contains", 1, "collections.contains"),
                ("Count", 0, "dotnet.observable_collection_count"),
                ("Clear", 0, "collections.clear"),
                ("IndexOf", 1, "collections.index_of"),
                ("Sort", 0, "dotnet.array_sort"),
                ("Reverse", 0, "collections.reverse"),
                ("ToArray", 0, "collections.clone"),
                ("Item", 1, "dotnet.list_get_checked"),
                ("Insert", 2, "collections.insert"),
                ("AddRange", 1, "dotnet.list_add_range"),
                ("InsertRange", 2, "collections.insert_range"),
                ("RemoveRange", 2, "collections.remove_range"),
                ("GetRange", 2, "collections.get_range"),
                ("SetRange", 2, "collections.set_range"),
                ("Exists", 1, "dotnet.array_exists"),
                ("Find", 1, "dotnet.array_find"),
                ("FindAll", 1, "dotnet.array_find_all"),
                ("BinarySearch", 1, "dotnet.array_binary_search"),
                ("EnsureCapacity", 1, "dotnet.list_ensure_capacity"),
                ("TrimExcess", 0, "dotnet.list_trim_excess"),
            ],
        ),
        // .NET `Dictionary<K,V>` is shape-identical to ECMA-262 §24.1
        // `Map`. The wrapper materializes a real `ObjectKind::Map` via
        // `ecma:map/new` and forwards every method to the corresponding
        // `Map.prototype.*` host fn. No `vybe:types` involvement.
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("Dictionary")
                .with_constructor(
                    ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")),
                )
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.dict_new_ignore_arg"),
                )
                .with_method(MethodDef::new(
                    "Add",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "set")),
                ))
                .with_method(MethodDef::new(
                    "TryAdd",
                    2,
                    MethodBody::Common("dotnet.dict_try_add".into()),
                ))
                .with_method(MethodDef::new(
                    "Item",
                    1,
                    MethodBody::Common("dotnet.dict_get_or_throw".into()),
                ))
                .with_method(MethodDef::new(
                    "ContainsKey",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "has")),
                ))
                .with_method(MethodDef::new(
                    "ContainsValue",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "containsValue")),
                ))
                .with_method(MethodDef::new(
                    "Remove",
                    1,
                    MethodBody::Common("dotnet.dict_remove".into()),
                ))
                .with_method(MethodDef::new(
                    "Keys",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "keys")),
                ))
                .with_method(MethodDef::new(
                    "Values",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "values")),
                ))
                .with_method(MethodDef::new(
                    "Entries",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "entries")),
                ))
                .with_method(MethodDef::new(
                    "EntriesSorted",
                    0,
                    MethodBody::Common("dotnet.sorted_dictionary_entries".into()),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "clear")),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "size")),
                ))
                .with_method(MethodDef::new(
                    "GetValueOrDefault",
                    1,
                    MethodBody::Common("dotnet.dict_get_value_or_default".into()),
                ))
                .with_method(MethodDef::new(
                    "TryGetValue",
                    2,
                    MethodBody::Common("dotnet.dict_try_get_value".into()),
                ))
                .with_method(MethodDef::new(
                    "EnsureCapacity",
                    1,
                    MethodBody::Common("dotnet.dict_ensure_capacity".into()),
                ))
                .with_method(MethodDef::new(
                    "TrimExcess",
                    0,
                    MethodBody::Common("dotnet.dict_trim_excess".into()),
                )),
        ),
        // .NET `Queue<T>` is a JS Array used FIFO - `Enqueue` appends
        // (push), `Dequeue` removes from the front (shift), `Peek`
        // looks at the front (`ecma:array.first`).
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("Queue")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new(
                    "Enqueue",
                    1,
                    MethodBody::Common("collections.push".into()),
                ))
                .with_method(MethodDef::new(
                    "Dequeue",
                    0,
                    MethodBody::Common("collections.shift".into()),
                ))
                .with_method(MethodDef::new(
                    "Peek",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "first")),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("collections.length".into()),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::Common("collections.clear".into()),
                ))
                .with_method(MethodDef::new(
                    "Contains",
                    1,
                    MethodBody::Common("collections.contains".into()),
                ))
                .with_method(MethodDef::new(
                    "ToArray",
                    0,
                    MethodBody::Common("collections.clone".into()),
                )),
        ),
        // .NET `Stack<T>` is a JS Array used LIFO - `Push` appends
        // (push), `Pop` removes from the end (pop), `Peek` looks at
        // the end (`ecma:array.last`).
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("Stack")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new(
                    "Push",
                    1,
                    MethodBody::Common("collections.push".into()),
                ))
                .with_method(MethodDef::new(
                    "Pop",
                    0,
                    MethodBody::Common("collections.pop".into()),
                ))
                .with_method(MethodDef::new(
                    "Peek",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:array", "last")),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("collections.length".into()),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::Common("collections.clear".into()),
                ))
                .with_method(MethodDef::new(
                    "Contains",
                    1,
                    MethodBody::Common("collections.contains".into()),
                ))
                .with_method(MethodDef::new(
                    "ToArray",
                    0,
                    MethodBody::Common("collections.clone".into()),
                )),
        ),
        // .NET `HashSet<T>` is a real ECMA-262 §24.2 `Set`. Constructor
        // creates an `ObjectKind::Set`; methods route through the
        // matching `ecma:set.*` host fns.
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("HashSet")
                .with_constructor(
                    ConstructorDef::new(0).with_backing(HostTarget::new("ecma:set", "new")),
                )
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.set_new_from_iterable"),
                )
                .with_method(MethodDef::new(
                    "Add",
                    1,
                    MethodBody::Common("dotnet.hashset_add".into()),
                ))
                .with_method(MethodDef::new(
                    "Remove",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "delete")),
                ))
                .with_method(MethodDef::new(
                    "Contains",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "has")),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "size")),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "clear")),
                ))
                .with_method(MethodDef::new(
                    "UnionWith",
                    1,
                    MethodBody::Common("dotnet.hashset_union_with".into()),
                ))
                .with_method(MethodDef::new(
                    "IntersectWith",
                    1,
                    MethodBody::Common("dotnet.hashset_intersect_with".into()),
                ))
                .with_method(MethodDef::new(
                    "ExceptWith",
                    1,
                    MethodBody::Common("dotnet.hashset_except_with".into()),
                ))
                .with_method(MethodDef::new(
                    "SymmetricExceptWith",
                    1,
                    MethodBody::Common("dotnet.hashset_symmetric_except_with".into()),
                ))
                .with_method(MethodDef::new(
                    "IsSubsetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_subset_of".into()),
                ))
                .with_method(MethodDef::new(
                    "IsSupersetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_superset_of".into()),
                ))
                .with_method(MethodDef::new(
                    "Overlaps",
                    1,
                    MethodBody::Common("dotnet.hashset_overlaps".into()),
                ))
                .with_method(MethodDef::new(
                    "SetEquals",
                    1,
                    MethodBody::Common("dotnet.hashset_set_equals".into()),
                ))
                .with_method(MethodDef::new(
                    "IsProperSubsetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_proper_subset_of".into()),
                ))
                .with_method(MethodDef::new(
                    "IsProperSupersetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_proper_superset_of".into()),
                )),
        ),
        // `ConcurrentDictionary` is a thread-safe `Dictionary` - same
        // shape (ECMA Map). Atomicity isn't modeled; methods route the
        // same way.
        DotnetClassExport::new(
            "dotnet.System.Collections.Concurrent",
            ClassType::new("ConcurrentDictionary")
                .with_constructor(
                    ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")),
                )
                .with_method(MethodDef::new(
                    "TryAdd",
                    2,
                    MethodBody::Common("dotnet.dict_try_add".into()),
                ))
                .with_method(MethodDef::new(
                    "TryGetValue",
                    2,
                    MethodBody::Common("dotnet.dict_try_get_value".into()),
                ))
                .with_method(MethodDef::new(
                    "AddOrUpdate",
                    3,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "set")),
                ))
                .with_method(MethodDef::new(
                    "GetOrAdd",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "get")),
                ))
                .with_method(MethodDef::new(
                    "ContainsKey",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "has")),
                ))
                .with_method(MethodDef::new(
                    "Remove",
                    1,
                    MethodBody::Common("dotnet.dict_remove".into()),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "clear")),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "size")),
                )),
        ),
        collection_class_common(
            "dotnet.System.Collections.Concurrent",
            "ConcurrentBag",
            "collections.new",
            &[
                ("Add", 1, "collections.push"),
                ("TryTake", 1, "dotnet.concurrent_stack_try_pop"),
                ("TryPeek", 1, "dotnet.concurrent_stack_try_peek"),
                ("Count", 0, "dotnet.observable_collection_count"),
            ],
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Concurrent",
            ClassType::new("BlockingCollection")
                .with_constructor(
                    ConstructorDef::new(0).with_common_backing("dotnet.blocking_collection_new"),
                )
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.blocking_collection_new"),
                )
                .with_method(MethodDef::new(
                    "Add",
                    1,
                    MethodBody::Common("dotnet.blocking_collection_add".into()),
                ))
                .with_method(MethodDef::new(
                    "TryAdd",
                    1,
                    MethodBody::Common("dotnet.blocking_collection_try_add".into()),
                ))
                .with_method(MethodDef::new(
                    "Take",
                    0,
                    MethodBody::Common("dotnet.blocking_collection_take".into()),
                ))
                .with_method(MethodDef::new(
                    "TryTake",
                    1,
                    MethodBody::Common("dotnet.blocking_collection_take".into()),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("dotnet.blocking_collection_count".into()),
                ))
                .with_method(MethodDef::new(
                    "CompleteAdding",
                    0,
                    MethodBody::Common("dotnet.blocking_collection_complete_adding".into()),
                ))
                .with_method(MethodDef::new(
                    "IsAddingCompleted",
                    0,
                    MethodBody::Common("dotnet.blocking_collection_is_completed".into()),
                ))
                .with_method(MethodDef::new(
                    "IsCompleted",
                    0,
                    MethodBody::Common("dotnet.blocking_collection_is_completed".into()),
                ))
                .with_method(MethodDef::new(
                    "GetConsumingEnumerable",
                    0,
                    MethodBody::Common("dotnet.blocking_collection_items".into()),
                )),
        ),
        // ConcurrentQueue / ConcurrentStack - same shape as their
        // non-concurrent counterparts (Array). Atomicity isn't
        // modeled at this layer.
        DotnetClassExport::new(
            "dotnet.System.Collections.Concurrent",
            ClassType::new("ConcurrentQueue")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new(
                    "Enqueue",
                    1,
                    MethodBody::Common("collections.push".into()),
                ))
                .with_method(MethodDef::new(
                    "TryDequeue",
                    1,
                    MethodBody::Common("dotnet.concurrent_queue_try_dequeue".into()),
                ))
                .with_method(MethodDef::new(
                    "TryPeek",
                    1,
                    MethodBody::Common("dotnet.concurrent_queue_try_peek".into()),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("collections.length".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Concurrent",
            ClassType::new("ConcurrentStack")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new(
                    "Push",
                    1,
                    MethodBody::Common("collections.push".into()),
                ))
                .with_method(MethodDef::new(
                    "TryPop",
                    0,
                    MethodBody::Common("collections.pop".into()),
                ))
                .with_method(MethodDef::new(
                    "TryPop",
                    1,
                    MethodBody::Common("dotnet.concurrent_stack_try_pop".into()),
                ))
                .with_method(MethodDef::new(
                    "TryPeek",
                    0,
                    MethodBody::Common("collections.get".into()),
                ))
                .with_method(MethodDef::new(
                    "TryPeek",
                    1,
                    MethodBody::Common("dotnet.concurrent_stack_try_peek".into()),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("collections.length".into()),
                )),
        ),
        // `SortedDictionary<K,V>` keeps the `ecma:map` backing (so lookup /
        // insert / removal stay O(1) and order-independent) and sorts its
        // Keys / Values / entry views at read time via the shared sorted core.
        // Natural ordering falls out of the shared functions' null-comparator
        // path, so no comparator needs to be stored.
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("SortedDictionary")
                .with_constructor(
                    ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")),
                )
                .with_method(MethodDef::new(
                    "Add",
                    2,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "set")),
                ))
                .with_method(MethodDef::new(
                    "Item",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "get")),
                ))
                .with_method(MethodDef::new(
                    "ContainsKey",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "has")),
                ))
                .with_method(MethodDef::new(
                    "Remove",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "delete")),
                ))
                .with_method(MethodDef::new(
                    "Keys",
                    0,
                    MethodBody::Common("dotnet.sorted_map_keys".into()),
                ))
                .with_method(MethodDef::new(
                    "Values",
                    0,
                    MethodBody::Common("dotnet.sorted_map_values".into()),
                ))
                .with_method(MethodDef::new(
                    "EntriesSorted",
                    0,
                    MethodBody::Common("dotnet.sorted_map_entries".into()),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "clear")),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:map", "size")),
                ))
                .with_method(MethodDef::new(
                    "GetValueOrDefault",
                    1,
                    MethodBody::Common("dotnet.dict_get_value_or_default".into()),
                ))
                .with_method(MethodDef::new(
                    "GetValueOrDefault",
                    2,
                    MethodBody::Common("dotnet.dict_get_value_or_default".into()),
                ))
                .with_method(MethodDef::new(
                    "TryGetValue",
                    2,
                    MethodBody::Common("dotnet.dict_try_get_value".into()),
                )),
        ),
        // `SortedSet<T>` keeps the `ecma:set` backing (so Add / Contains /
        // Remove / Count and the whole set-algebra surface reuse the same host
        // set ops as `HashSet`). Only the ordered reads are adapted: `foreach`
        // is rewritten to `ElementsSorted()` and `GetViewBetween` spreads the
        // set to a sorted array, both via the shared sorted core, then the view
        // is rebuilt as a set so its own methods resolve through the set path.
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("SortedSet")
                .with_constructor(
                    ConstructorDef::new(0).with_backing(HostTarget::new("ecma:set", "new")),
                )
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.set_new_ignore_comparer"),
                )
                .with_method(MethodDef::new(
                    "Add",
                    1,
                    MethodBody::Common("dotnet.hashset_add".into()),
                ))
                .with_method(MethodDef::new(
                    "Remove",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "delete")),
                ))
                .with_method(MethodDef::new(
                    "Contains",
                    1,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "has")),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "size")),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::HostCall(HostTarget::new("ecma:set", "clear")),
                ))
                .with_method(MethodDef::new(
                    "UnionWith",
                    1,
                    MethodBody::Common("dotnet.hashset_union_with".into()),
                ))
                .with_method(MethodDef::new(
                    "IntersectWith",
                    1,
                    MethodBody::Common("dotnet.hashset_intersect_with".into()),
                ))
                .with_method(MethodDef::new(
                    "ExceptWith",
                    1,
                    MethodBody::Common("dotnet.hashset_except_with".into()),
                ))
                .with_method(MethodDef::new(
                    "SymmetricExceptWith",
                    1,
                    MethodBody::Common("dotnet.hashset_symmetric_except_with".into()),
                ))
                .with_method(MethodDef::new(
                    "IsSubsetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_subset_of".into()),
                ))
                .with_method(MethodDef::new(
                    "IsSupersetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_superset_of".into()),
                ))
                .with_method(MethodDef::new(
                    "Overlaps",
                    1,
                    MethodBody::Common("dotnet.hashset_overlaps".into()),
                ))
                .with_method(MethodDef::new(
                    "SetEquals",
                    1,
                    MethodBody::Common("dotnet.hashset_set_equals".into()),
                ))
                .with_method(MethodDef::new(
                    "IsProperSubsetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_proper_subset_of".into()),
                ))
                .with_method(MethodDef::new(
                    "IsProperSupersetOf",
                    1,
                    MethodBody::Common("dotnet.hashset_is_proper_superset_of".into()),
                ))
                .with_method(MethodDef::new(
                    "Min",
                    0,
                    MethodBody::Common("dotnet.sorted_set_min".into()),
                ))
                .with_method(MethodDef::new(
                    "Max",
                    0,
                    MethodBody::Common("dotnet.sorted_set_max".into()),
                ))
                .with_method(MethodDef::new(
                    "ElementsSorted",
                    0,
                    MethodBody::Common("dotnet.sorted_set_elements".into()),
                ))
                .with_method(MethodDef::new(
                    "GetViewBetween",
                    2,
                    MethodBody::Common("dotnet.sorted_set_view_between".into()),
                )),
        ),
        constructor_class(
            "dotnet.System.Collections.Generic",
            "SortedList",
            "ecma:map",
            "new",
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Generic",
            ClassType::new("LinkedList")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new(
                    "AddFirst",
                    1,
                    MethodBody::Common("dotnet.linked_list_add_first".into()),
                ))
                .with_method(MethodDef::new(
                    "AddLast",
                    1,
                    MethodBody::Common("dotnet.linked_list_add_last".into()),
                ))
                .with_method(MethodDef::new(
                    "Find",
                    1,
                    MethodBody::Common("dotnet.linked_list_find".into()),
                ))
                .with_method(MethodDef::new(
                    "First",
                    0,
                    MethodBody::Common("dotnet.linked_list_first".into()),
                ))
                .with_method(MethodDef::new(
                    "InsertAtRaw",
                    2,
                    MethodBody::Common("collections.insert".into()),
                ))
                .with_method(MethodDef::new(
                    "RemoveAt",
                    1,
                    MethodBody::Common("collections.remove_at".into()),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::Common("collections.clear".into()),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("collections.length".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections",
            ClassType::new("ArrayList")
                .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                .with_method(MethodDef::new(
                    "Add",
                    1,
                    MethodBody::Common("dotnet.list_add".into()),
                ))
                .with_method(MethodDef::new(
                    "Remove",
                    1,
                    MethodBody::Common("collections.remove".into()),
                ))
                .with_method(MethodDef::new(
                    "RemoveAt",
                    1,
                    MethodBody::Common("collections.remove_at".into()),
                ))
                .with_method(MethodDef::new(
                    "Contains",
                    1,
                    MethodBody::Common("collections.contains".into()),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("collections.length".into()),
                ))
                .with_method(MethodDef::new(
                    "Capacity",
                    0,
                    MethodBody::Common("dotnet.list_capacity".into()),
                ))
                .with_method(MethodDef::new(
                    "EnsureCapacity",
                    1,
                    MethodBody::Common("dotnet.list_ensure_capacity".into()),
                ))
                .with_method(MethodDef::new(
                    "TrimExcess",
                    0,
                    MethodBody::Common("dotnet.list_trim_excess".into()),
                ))
                .with_method(MethodDef::new(
                    "Clear",
                    0,
                    MethodBody::Common("collections.clear".into()),
                ))
                .with_method(MethodDef::new(
                    "IndexOf",
                    1,
                    MethodBody::Common("collections.index_of".into()),
                ))
                .with_method(MethodDef::new(
                    "IndexOf",
                    2,
                    MethodBody::Common("collections.index_of_from".into()),
                ))
                .with_method(MethodDef::new(
                    "LastIndexOf",
                    1,
                    MethodBody::Common("collections.last_index_of".into()),
                ))
                .with_method(MethodDef::new(
                    "LastIndexOf",
                    2,
                    MethodBody::Common("collections.last_index_of_from".into()),
                ))
                .with_method(MethodDef::new(
                    "Sort",
                    0,
                    MethodBody::Common("dotnet.array_sort".into()),
                ))
                .with_method(MethodDef::new(
                    "Reverse",
                    0,
                    MethodBody::Common("collections.reverse".into()),
                ))
                .with_method(MethodDef::new(
                    "Reverse",
                    2,
                    MethodBody::Common("collections.reverse_range".into()),
                ))
                .with_method(MethodDef::new(
                    "ToArray",
                    0,
                    MethodBody::Common("collections.clone".into()),
                ))
                .with_method(MethodDef::new(
                    "Clone",
                    0,
                    MethodBody::Common("collections.clone".into()),
                ))
                .with_method(MethodDef::new(
                    "Item",
                    1,
                    MethodBody::Common("collections.get".into()),
                ))
                .with_method(MethodDef::new(
                    "Insert",
                    2,
                    MethodBody::Common("collections.insert".into()),
                ))
                .with_method(MethodDef::new(
                    "InsertRange",
                    2,
                    MethodBody::Common("collections.insert_range".into()),
                ))
                .with_method(MethodDef::new(
                    "RemoveRange",
                    2,
                    MethodBody::Common("collections.remove_range".into()),
                ))
                .with_method(MethodDef::new(
                    "GetRange",
                    2,
                    MethodBody::Common("collections.get_range".into()),
                ))
                .with_method(MethodDef::new(
                    "SetRange",
                    2,
                    MethodBody::Common("collections.set_range".into()),
                ))
                .with_method(MethodDef::new(
                    "BinarySearch",
                    1,
                    MethodBody::Common("dotnet.array_binary_search".into()),
                ))
                .with_method(MethodDef::new(
                    "AddRange",
                    1,
                    MethodBody::Common("collections.concat".into()),
                )),
        ),
        constructor_class("dotnet.System.Collections", "Hashtable", "ecma:map", "new"),
        collection_class_common(
            "dotnet.System.Collections.ObjectModel",
            "ObservableCollection",
            "collections.new",
            &[
                ("Add", 1, "dotnet.observable_collection_add"),
                ("Remove", 1, "dotnet.observable_collection_remove"),
                ("RemoveAt", 1, "dotnet.observable_collection_remove_at"),
                ("Insert", 2, "dotnet.observable_collection_insert"),
                ("Move", 2, "dotnet.observable_collection_move"),
                ("Clear", 0, "dotnet.observable_collection_clear"),
                ("Count", 0, "collections.length"),
                ("Item", 1, "dotnet.list_get_checked"),
                ("ToArray", 0, "collections.clone"),
                ("Items", 0, "dotnet.observable_collection_items"),
                (
                    "OnCollectionChanged",
                    1,
                    "dotnet.observable_collection_on_changed",
                ),
            ],
        ),
        collection_class_common(
            "dotnet.System.Collections.ObjectModel",
            "ReadOnlyObservableCollection",
            "dotnet.readonly_observable_collection_new",
            &[
                ("Count", 0, "collections.length"),
                ("Item", 1, "dotnet.list_get_checked"),
                ("ToArray", 0, "collections.clone"),
            ],
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections.Specialized",
            ClassType::new("NotifyCollectionChangedEventArgs").with_constructor(
                ConstructorDef::new(1)
                    .with_common_backing("dotnet.notify_collection_changed_event_args_new"),
            ),
        ),
        DotnetClassExport::new(
            "dotnet.System.ComponentModel",
            ClassType::new("PropertyChangedEventArgs").with_constructor(
                ConstructorDef::new(1)
                    .with_common_backing("dotnet.property_changed_event_args_new"),
            ),
        ),
        DotnetClassExport::new(
            "dotnet.System.Collections",
            ClassType::new("Collection")
                .with_constructor(
                    ConstructorDef::new(0).with_common_backing("dotnet.vb_collection_new"),
                )
                .with_method(MethodDef::new(
                    "Add",
                    1,
                    MethodBody::Common("dotnet.vb_collection_add".into()),
                ))
                .with_method(MethodDef::new(
                    "Add",
                    2,
                    MethodBody::Common("dotnet.vb_collection_add".into()),
                ))
                .with_method(MethodDef::new(
                    "Item",
                    1,
                    MethodBody::Common("dotnet.vb_collection_item".into()),
                ))
                .with_method(MethodDef::new(
                    "Count",
                    0,
                    MethodBody::Common("dotnet.vb_collection_count".into()),
                ))
                .with_method(MethodDef::new(
                    "ToArray",
                    0,
                    MethodBody::Common("dotnet.vb_collection_to_array".into()),
                ))
                .with_method(MethodDef::new(
                    "Contains",
                    1,
                    MethodBody::Common("dotnet.vb_collection_contains".into()),
                ))
                .with_method(MethodDef::new(
                    "Remove",
                    1,
                    MethodBody::Common("dotnet.vb_collection_remove".into()),
                )),
        ),
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
    if name == "List" {
        class = class.with_constructor(ConstructorDef::new(1).with_common_backing(ctor_common));
    } else if name == "ObservableCollection" {
        class = class.with_constructor(
            ConstructorDef::new(1).with_common_backing("dotnet.list_new_from_iterable"),
        );
    } else if name == "ReadOnlyObservableCollection" {
        class = class.with_constructor(
            ConstructorDef::new(1).with_common_backing("dotnet.readonly_observable_collection_new"),
        );
    }
    for (method, arity, common) in methods {
        class = class.with_method(MethodDef::new(
            *method,
            *arity,
            MethodBody::Common((*common).into()),
        ));
    }
    DotnetClassExport::new(interface, class)
}
