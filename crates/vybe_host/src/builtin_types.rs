//! Register built-in types in the VM's TypeRegistry.
//! Each type gets a vtable with methods resolved via the host function registry.
//! This replaces the legacy type_methods table with proper WASM GC-style dispatch.

use vybe_bytecode::{VM, TypeDef, Method, Value};

pub fn register_all(vm: &mut VM) {
    // Helper: look up host fn index by (module, name)
    let h = |vm: &VM, module: &str, name: &str| -> Option<usize> {
        vm.host_registry.get(&(module.to_string(), name.to_string())).copied()
    };

    // --- Object (type 0, already created) ---
    // Add universal methods to Object
    if let Some(idx) = h(vm, "vybe:convert", "toString") {
        vm.type_registry.add_host_method(0, "tostring", idx);
    }
    if let Some(idx) = h(vm, "vybe:convert", "typeName") {
        vm.type_registry.add_host_method(0, "gettype", idx);
    }
    if let Some(idx) = h(vm, "vybe:convert", "toString") {
        vm.type_registry.add_host_method(0, "gethashcode", idx);
        vm.type_registry.add_host_method(0, "equals", idx);
    }

    // --- String ---
    //
    // VB/C#/.NET-style instance method names (Contains, IndexOf,
    // Substring, etc.) bound directly to ECMA-262 §22.1
    // String.prototype. Same import surface JS engines satisfy
    // natively. Methods needing arg adaptation (.NET's `Insert`,
    // `Remove`, `LastIndexOf` with different-shape args) are still
    // bound to vybe:string for the legacy impl until language stdlib
    // adapter chunks are written.
    let _string_id = {
        let mut t = TypeDef::new("String");
        for (method, module, fname) in &[
            ("contains", "ecma:string", "includes"),
            ("toupper", "ecma:string", "toUpperCase"),
            ("tolower", "ecma:string", "toLowerCase"),
            ("trim", "ecma:string", "trim"),
            ("trimstart", "ecma:string", "trimStart"),
            ("trimend", "ecma:string", "trimEnd"),
            ("startswith", "ecma:string", "startsWith"),
            ("endswith", "ecma:string", "endsWith"),
            ("indexof", "ecma:string", "indexOf"),
            ("lastindexof", "ecma:string", "lastIndexOf"),
            ("substring", "ecma:string", "substring"),
            ("replace", "ecma:string", "replace"),
            ("split", "ecma:string", "split"),
            ("padleft", "ecma:string", "padStart"),
            ("padright", "ecma:string", "padEnd"),
            ("tostring", "vybe:convert", "toString"),
            ("toupperinvariant", "ecma:string", "toUpperCase"),
            ("tolowerinvariant", "ecma:string", "toLowerCase"),
            ("chars", "ecma:string", "charAt"),
            ("insert", "vybe:string", "mid"),
            ("remove", "ecma:string", "slice"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0); // inherits from Object
        vm.type_registry.register(t)
    };

    // --- List / Array ---
    //
    // Phase 7b: List<T> is a JS Array per ECMA-262 §23.1. Every method
    // routes through `ecma:array/*` — the same host fns v8 / other
    // runtimes expose via the wasm-js-builtins proposal.
    //
    // `list.Count` and `list.Length` are .NET/JS property-style reads;
    // the compile path (struct_get "count") auto-invokes `__get_count`
    // installed on each List instance by `vybe:types/listNew`. The
    // method-call form `list.Count()` dispatches via TypeRegistry to
    // `ecma:array/length` which returns an i32.
    //
    // Transitional entries flagged `vybe:types/list*` remain only where
    // a direct `ecma:array` equivalent doesn't exist (RemoveAt,
    // AddRange semantic mismatch with concat, BinarySearch, GetRange /
    // SetRange, TryPeek/TryPop at-end). Each lands on a JS-shape
    // primitive over time.
    let list_id = {
        let mut t = TypeDef::new("List");
        if let Some(idx) = h(vm, "vybe:types", "listNew") {
            t.constructor = Some(Method::HostFn(idx));
        }
        for (method, module, fname) in &[
            ("add", "ecma:array", "push"),
            ("remove", "vybe:types", "listRemove"),       // by-value; no ecma:array direct
            ("removeat", "vybe:types", "listRemoveAt"),   // splice pattern; to migrate
            ("contains", "ecma:array", "includes"),
            ("count", "ecma:array", "length"),
            ("clear", "vybe:types", "listClear"),          // setLength(0) pattern; to migrate
            ("indexof", "ecma:array", "indexOf"),
            ("sort", "ecma:array", "sort"),
            ("reverse", "ecma:array", "reverse"),
            ("toarray", "ecma:array", "slice"),
            ("item", "ecma:array", "get"),
            ("lastindexof", "ecma:array", "lastIndexOf"),
            ("insert", "vybe:types", "listInsert"),        // splice(idx, 0, v); to migrate
            ("addrange", "vybe:types", "listAddRange"),    // in-place extend; concat returns new
            ("capacity", "ecma:array", "length"),
            ("insertrange", "vybe:types", "listInsertRange"),
            ("removerange", "vybe:types", "listRemoveRange"),
            ("getrange", "vybe:types", "listGetRange"),
            ("setrange", "vybe:types", "listSetRange"),
            ("binarysearch", "vybe:types", "listBinarySearch"),
            ("clone", "ecma:array", "slice"),
            ("copyto", "ecma:array", "slice"),
            ("trimtosize", "ecma:array", "length"),
            ("enqueue", "ecma:array", "push"),
            ("trydequeue", "ecma:array", "shift"),
            ("trypop", "ecma:array", "pop"),
            ("trypeek", "vybe:types", "listLast"),         // get(length-1); compound
            // JS Array methods — direct pass-through.
            ("push", "ecma:array", "push"),
            ("pop", "ecma:array", "pop"),
            ("shift", "ecma:array", "shift"),
            ("unshift", "ecma:array", "unshift"),
            ("join", "ecma:array", "join"),
            ("includes", "ecma:array", "includes"),
            ("slice", "ecma:array", "slice"),
            ("concat", "ecma:array", "concat"),
            ("splice", "ecma:array", "splice"),
            ("fill", "ecma:array", "fill"),
            ("flat", "ecma:array", "flat"),
            ("find", "ecma:array", "find"),
            ("findindex", "ecma:array", "findIndex"),
            ("keys", "ecma:array", "keys"),
            ("values", "ecma:array", "values"),
            ("at", "ecma:array", "at"),
            ("copywithin", "ecma:array", "copyWithin"),
            ("entries", "ecma:array", "entries"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t)
    };

    // Also register "ArrayList" and "Array" as aliases for List
    let _ = vm.type_registry.register(TypeDef::new("ArrayList").with_parent(list_id));
    let _ = vm.type_registry.register(TypeDef::new("Array").with_parent(list_id));

    // --- Dictionary ---
    //
    // Phase 7b: Dictionary is a plain JS Object per ECMA-262 §19.1 —
    // `Dictionary<string, T>` is shape-identical to an object used as
    // a string-keyed map. Methods route through `ecma:object/*`
    // host fns (working directly on the object's own properties).
    //
    // Transition note: VB/C# `d.Count` is a property read (not a
    // method call). The old `vybe:types/dictAdd` kept a `count` field
    // on the dict updated on every write so `struct_get` returned a
    // number; `ecma:object/length` is a callable returning a count
    // when invoked. A later step will install `__get_count` auto-getters
    // on construction so `d.Count` property reads auto-invoke the
    // length call — until then, `.Count` reads expose the function
    // value. Code using explicit `d.Count()` works today; bare `d.Count`
    // is a pending follow-up. Known regression count tracked in the
    // phase notes.
    let dict_id = {
        let mut t = TypeDef::new("Dictionary");
        for (method, module, fname) in &[
            ("add", "ecma:object", "set"),
            ("item", "ecma:object", "get"),
            // `hasOwn` (own-only, returns Value::Bool) is what VB/C#
            // `ContainsKey` expects — string-coerces to "true"/"false".
            // `has` would return I32 ("1"/"0") which breaks VB prints.
            ("containskey", "ecma:object", "hasOwn"),
            ("containsvalue", "ecma:object", "hasOwn"),
            ("remove", "ecma:object", "delete"),
            ("keys", "ecma:object", "keys"),
            ("values", "ecma:object", "values"),
            ("clear", "vybe:types", "dictClear"),
            ("count", "ecma:object", "length"),
            ("trygetvalue", "ecma:object", "get"),
            ("tryadd", "ecma:object", "set"),
            ("addorupdate", "ecma:object", "set"),
            ("getoradd", "ecma:object", "get"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t)
    };
    let _ = vm.type_registry.register(TypeDef::new("Hashtable").with_parent(dict_id));

    // --- Queue ---
    //
    // Phase 7b: `.NET Queue<T>` is a JS Array used FIFO — `Enqueue` →
    // `push` appends at the end, `Dequeue` → `shift` removes from the
    // front. Property-style `q.Count` works via the __get_count
    // auto-getter installed by `vybe:types/queueNew`.
    {
        let mut t = TypeDef::new("Queue");
        for (method, module, fname) in &[
            ("enqueue", "ecma:array", "push"),
            ("dequeue", "ecma:array", "shift"),
            ("peek", "vybe:types", "queuePeek"),     // get(0); compound
            ("count", "ecma:array", "length"),
            ("clear", "vybe:types", "listClear"),    // setLength(0); to migrate
            ("contains", "ecma:array", "includes"),
            ("toarray", "ecma:array", "slice"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Stack ---
    //
    // Phase 7b: `.NET Stack<T>` is a JS Array used LIFO — `Push` →
    // `push` appends at the end, `Pop` → `pop` removes from the end.
    {
        let mut t = TypeDef::new("Stack");
        for (method, module, fname) in &[
            ("push", "ecma:array", "push"),
            ("pop", "ecma:array", "pop"),
            ("peek", "vybe:types", "stackPeek"),     // get(length-1); compound
            ("count", "ecma:array", "length"),
            ("clear", "vybe:types", "listClear"),
            ("contains", "ecma:array", "includes"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- HashSet ---
    //
    // Phase 7b: `.NET HashSet<T>` is a JS Set per ECMA-262 §24.2.
    {
        let mut t = TypeDef::new("HashSet");
        for (method, module, fname) in &[
            ("add", "ecma:set", "add"),
            ("contains", "ecma:set", "has"),
            ("remove", "ecma:set", "delete"),
            ("count", "ecma:set", "size"),
            ("clear", "ecma:set", "clear"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- StringBuilder ---
    {
        let mut t = TypeDef::new("StringBuilder");
        for (method, fname) in &[
            ("append", "sbAppend"), ("appendline", "sbAppendLine"), ("appendformat", "sbAppend"),
            ("tostring", "sbToString"), ("clear", "sbClear"),
            ("insert", "sbInsert"), ("replace", "sbReplace"),
            ("chars", "sbToString"), ("equals", "sbToString"),
            ("remove", "sbClear"), ("length", "sbToString"),
            ("ensurecapacity", "sbToString"), ("capacity", "sbToString"),
            ("copyto", "sbToString"), ("gettype", "sbToString"),
        ] {
            if let Some(idx) = h(vm, "vybe:types", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- DateTime ---
    {
        let mut t = TypeDef::new("DateTime");
        for (method, fname) in &[
            ("adddays", "dateTimeAddDays"), ("addhours", "dateTimeAddHours"),
            ("addminutes", "dateTimeAddMinutes"), ("addseconds", "dateTimeAddSeconds"),
            ("addmonths", "dateTimeAddMonths"), ("addyears", "dateTimeAddYears"),
            ("tostring", "dateTimeToString"), ("toshortdatestring", "dateTimeToShortDate"),
            ("tolongdatestring", "dateTimeToString"),
        ] {
            if let Some(idx) = h(vm, "vybe:types", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- SqlConnection ---
    {
        let mut t = TypeDef::new("SqlConnection");
        for (method, fname) in &[
            ("open", "open"), ("close", "close"),
            ("createcommand", "createCommand"),
            ("begintransaction", "beginTransaction"),
        ] {
            if let Some(idx) = h(vm, "vybe:database", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- SqlCommand ---
    {
        let mut t = TypeDef::new("SqlCommand");
        for (method, fname) in &[
            ("executenonquery", "execute"), ("executescalar", "scalar"),
            ("executereader", "query"),
        ] {
            if let Some(idx) = h(vm, "vybe:database", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- TcpClient ---
    {
        let mut t = TypeDef::new("TcpClient");
        for (method, fname) in &[
            ("close", "tcpClientClose"), ("getstream", "tcpClientGetStream"),
        ] {
            if let Some(idx) = h(vm, "dotnet:sockets", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- TcpListener ---
    {
        let mut t = TypeDef::new("TcpListener");
        for (method, fname) in &[
            ("start", "tcpListenerStart"), ("stop", "tcpListenerStop"),
            ("accepttcpclient", "tcpListenerAccept"), ("pending", "tcpListenerPending"),
        ] {
            if let Some(idx) = h(vm, "dotnet:sockets", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- UdpClient ---
    {
        let mut t = TypeDef::new("UdpClient");
        for (method, fname) in &[
            ("send", "udpClientSend"), ("receive", "udpClientReceive"), ("close", "udpClientClose"),
        ] {
            if let Some(idx) = h(vm, "dotnet:sockets", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- StreamReader ---
    {
        let mut t = TypeDef::new("StreamReader");
        for (method, fname) in &[
            ("readline", "streamReaderReadLine"), ("readtoend", "streamReaderReadToEnd"),
            ("close", "streamWriterClose"),
        ] {
            if let Some(idx) = h(vm, "dotnet:io", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- StreamWriter ---
    {
        let mut t = TypeDef::new("StreamWriter");
        for (method, fname) in &[
            ("writeline", "streamWriterWriteLine"), ("write", "streamWriterWrite"),
            ("flush", "streamWriterFlush"), ("close", "streamWriterClose"),
        ] {
            if let Some(idx) = h(vm, "dotnet:io", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Stopwatch ---
    {
        let mut t = TypeDef::new("Stopwatch");
        if let Some(idx) = h(vm, "wasi:clocks", "stopwatchStart") {
            t.methods.insert("start".into(), Method::HostFn(idx));
        }
        if let Some(idx) = h(vm, "wasi:clocks", "stopwatchStop") {
            t.methods.insert("stop".into(), Method::HostFn(idx));
        }
        if let Some(idx) = h(vm, "wasi:clocks", "stopwatchElapsed") {
            t.methods.insert("elapsedmilliseconds".into(), Method::HostFn(idx));
            t.methods.insert("elapsed".into(), Method::HostFn(idx));
        }
        if let Some(idx) = h(vm, "wasi:clocks", "stopwatchReset") {
            t.methods.insert("reset".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Random ---
    // .NET / VB / Dart `System.Random` class. Methods bind directly to real
    // WASI random primitives (`wasi:random/insecure/get-insecure-random-u64`).
    // Range-bounded `.Next(min, max)` and `.NextDouble()` semantics (modulo
    // and [0,1) float conversion) are compiler-level lowerings — the raw
    // WASI u64 is what comes back here. Per-instance seeding is NOT
    // supported because WASI's insecure RNG is process-global; `new
    // Random(seed)` accepts a seed argument and ignores it.
    {
        let mut t = TypeDef::new("Random");
        if let Some(idx) = h(vm, "wasi:random/insecure", "get-insecure-random-u64") {
            t.methods.insert("next".into(), Method::HostFn(idx));
            t.methods.insert("nextdouble".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- XDocument ---
    {
        let mut t = TypeDef::new("XDocument");
        if let Some(idx) = h(vm, "vybe:xml", "toString") {
            t.methods.insert("tostring".into(), Method::HostFn(idx));
            t.methods.insert("save".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- XElement ---
    {
        let mut t = TypeDef::new("XElement");
        if let Some(idx) = h(vm, "vybe:xml", "toString") {
            t.methods.insert("tostring".into(), Method::HostFn(idx));
            t.methods.insert("value".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- DataTable ---
    {
        let mut t = TypeDef::new("DataTable");
        if let Some(idx) = h(vm, "vybe:data", "dataTableNewRow") {
            t.methods.insert("newrow".into(), Method::HostFn(idx));
        }
        if let Some(idx) = h(vm, "vybe:data", "dataTableAddRow") {
            t.methods.insert("rows".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Map (JS collection) ---
    //
    // Bound directly to `ecma:map.*` — same imports JS engines satisfy
    // natively, no Vybe-private shape required. Backing store is
    // `ObjectKind::Map(IndexMap<Value,Value>)` (SameValueZero keys per
    // ECMA-262 §24.1).
    {
        let mut t = TypeDef::new("Map");
        for (method, fname) in &[
            ("set", "set"), ("get", "get"), ("has", "has"),
            ("delete", "delete"), ("keys", "keys"), ("values", "values"),
            ("clear", "clear"), ("entries", "entries"), ("forEach", "forEach"),
        ] {
            if let Some(idx) = h(vm, "ecma:map", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        // `size` is a property updated by every mutator, but expose the
        // method form too for VB-style `.Count` callers and JS code that
        // does `m.size()` even though spec says it's a getter.
        if let Some(idx) = h(vm, "ecma:map", "size") {
            t.methods.insert("size".to_string(), Method::HostFn(idx));
            t.methods.insert("count".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Set (JS collection) ---
    //
    // Bound to `ecma:set.*`. Backing store is
    // `ObjectKind::Set(IndexSet<Value>)`.
    {
        let mut t = TypeDef::new("Set");
        for (method, fname) in &[
            ("add", "add"), ("has", "has"), ("delete", "delete"),
            ("values", "values"), ("keys", "values"), ("clear", "clear"),
            ("entries", "entries"), ("forEach", "forEach"),
        ] {
            if let Some(idx) = h(vm, "ecma:set", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        if let Some(idx) = h(vm, "ecma:set", "size") {
            t.methods.insert("size".to_string(), Method::HostFn(idx));
            t.methods.insert("count".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- WeakMap (JS collection) ---
    //
    // Bound to `ecma:weakmap.*` — share-nothing with Map (different
    // backing shape: parallel keys/values arrays per ES spec
    // garbage-collection model).
    {
        let mut t = TypeDef::new("WeakMap");
        for (method, fname) in &[
            ("set", "set"), ("get", "get"), ("has", "has"), ("delete", "delete"),
        ] {
            if let Some(idx) = h(vm, "ecma:weakmap", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- WeakSet (JS collection) ---
    {
        let mut t = TypeDef::new("WeakSet");
        for (method, fname) in &[
            ("add", "add"), ("has", "has"), ("delete", "delete"),
        ] {
            if let Some(idx) = h(vm, "ecma:weakset", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- RegExp (ECMA-262 §22.2) ---
    //
    // Bound to `ecma:regexp.*`. Backing object is a plain Object with
    // `source`/`flags`/etc. as own properties — re-compiled at every
    // call (regex crate's `Regex::new` is cheap for small patterns;
    // caching can come later).
    {
        let mut t = TypeDef::new("RegExp");
        for (method, fname) in &[
            ("test", "test"), ("exec", "exec"), ("toString", "toString"),
        ] {
            if let Some(idx) = h(vm, "ecma:regexp", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Intl.* (ECMA-402) ---
    //
    // Each Intl class is bound to its `ecma:intl/<class>` host module.
    // Constructors (via the namespace registry at `namespaces/intl.rs`)
    // stamp `__type` matching the class name so this registry resolves
    // `instance.method(...)` to the correct host fn.
    {
        let mut t = TypeDef::new("Collator");
        for (method, fname) in &[
            ("compare", "compare"), ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/collator", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("NumberFormat");
        for (method, fname) in &[
            ("format", "format"), ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/numberformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("DateTimeFormat");
        for (method, fname) in &[
            ("format", "format"), ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/datetimeformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("ListFormat");
        for (method, fname) in &[
            ("format", "format"), ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/listformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("PluralRules");
        for (method, fname) in &[
            ("select", "select"), ("selectRange", "selectRange"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/pluralrules", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("RelativeTimeFormat");
        for (method, fname) in &[
            ("format", "format"), ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/relativetimeformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("Segmenter");
        for (method, fname) in &[
            ("segment", "segment"), ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/segmenter", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("Locale");
        for (method, fname) in &[
            ("toString", "toString"), ("maximize", "maximize"),
            ("minimize", "minimize"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/locale", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("DisplayNames");
        for (method, fname) in &[
            ("of", "of"), ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/displaynames", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("DurationFormat");
        for (method, fname) in &[
            ("format", "format"), ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = h(vm, "ecma:intl/durationformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- TimeSpan ---
    {
        let t = TypeDef::new("TimeSpan");
        // Properties are stored on the object directly (totalseconds, etc.)
        vm.type_registry.register(t);
    }

    // --- Task ---
    // `Task.Start()` is a no-op: compiler emits Op::THREAD_SPAWN which
    // creates AND starts the continuation atomically (per WASM stack-
    // switching semantics). No host method needed.
    {
        let t = TypeDef::new("Task");
        vm.type_registry.register(t);
    }

    // --- Process ---
    {
        let mut t = TypeDef::new("Process");
        if let Some(idx) = h(vm, "vybe:types", "processWaitForExit") {
            t.methods.insert("waitforexit".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Promise (JS) ---
    {
        let t = TypeDef::new("Promise");
        vm.type_registry.register(t);
    }

    // ============================================================
    // GUI Control type hierarchy
    // ============================================================
    // Control is the abstract base for all UI controls
    let control_id = {
        let mut t = TypeDef::new("Control");
        // Common control methods — resolved via gui host functions
        for (method, module, fname) in &[
            ("show", "vybe:gui", "__ctrl_show"),
            ("close", "vybe:gui", "__ctrl_close"),
            ("focus", "vybe:gui", "__ctrl_focus"),
            ("hide", "vybe:gui", "__ctrl_hide"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0); // inherits from Object
        vm.type_registry.register(t)
    };

    // Register all concrete control types as subtypes of Control
    let control_type_names = [
        "Button", "Label", "TextBox", "CheckBox", "RadioButton", "ComboBox",
        "ListBox", "Panel", "GroupBox", "TabControl", "TabPage", "DataGridView",
        "ProgressBar", "TrackBar", "NumericUpDown", "DateTimePicker", "RichTextBox",
        "PictureBox", "MenuStrip", "ToolStrip", "StatusStrip", "SplitContainer",
        "FlowLayoutPanel", "TableLayoutPanel", "LinkLabel", "MaskedTextBox",
        "ListView", "WebBrowser", "MonthCalendar", "ContextMenuStrip",
        "Timer", "BindingSource", "DataSet", "ImageList", "ToolTip",
        "NotifyIcon", "ErrorProvider", "HelpProvider", "BackgroundWorker",
        "TreeView",
    ];
    for ct in &control_type_names {
        vm.type_registry.register(TypeDef::new(ct).with_parent(control_id));
    }
    // Form is special — inherits from Control, adds its own methods
    {
        let mut t = TypeDef::new("Form");
        for (method, module, fname) in &[
            ("show", "vybe:gui", "__ctrl_show"),
            ("close", "vybe:gui", "__ctrl_close"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(control_id);
        vm.type_registry.register(t);
    }

    // ============================================================
    // Register constructors for all types
    // ============================================================
    let ctor_mappings: &[(&str, &str, &str)] = &[
        ("Dictionary", "vybe:types", "dictNew"),
        ("Queue", "vybe:types", "queueNew"),
        ("Stack", "vybe:types", "stackNew"),
        ("HashSet", "vybe:types", "hashSetNew"),
        ("StringBuilder", "vybe:types", "stringBuilderNew"),
        ("DateTime", "vybe:types", "dateTimeNew"),
        ("SqlConnection", "vybe:database", "connect"),
        ("TcpClient", "dotnet:sockets", "tcpClientNew"),
        ("TcpListener", "dotnet:sockets", "tcpListenerNew"),
        ("UdpClient", "dotnet:sockets", "udpClientNew"),
        ("StreamReader", "dotnet:io", "streamReaderNew"),
        ("StreamWriter", "dotnet:io", "streamWriterNew"),
        ("Stopwatch", "wasi:clocks", "stopwatchNew"),
        // Random ctor: real WASI entropy. Seed argument (if any) is ignored —
        // wasi:random/insecure is a process-global PRNG with no per-instance
        // seed. The returned u64 becomes the "Random" receiver object; VB/C#
        // code treats it as opaque and calls methods on it. When compiler
        // lowering for range math lands, this becomes a trivial marker and
        // the method calls inline the WASI call themselves.
        ("Random", "wasi:random/insecure", "get-insecure-random-u64"),
        ("DataTable", "vybe:data", "dataTableNew"),
        ("DataSet", "vybe:data", "dataSetNew"),
        ("Point", "vybe:drawing", "pointNew"),
        ("Size", "vybe:drawing", "sizeNew"),
        ("Font", "vybe:drawing", "fontNew"),
    ];
    for (type_name, module, fname) in ctor_mappings {
        if let (Some(tid), Some(idx)) = (vm.type_registry.get_id(type_name), h(vm, module, fname)) {
            vm.type_registry.set_constructor(tid, Method::HostFn(idx));
        }
    }

    // Register GUI control constructors (new_Button, new_TextBox, etc.)
    let gui_ctors = [
        "Button", "Label", "TextBox", "CheckBox", "RadioButton", "ComboBox",
        "ListBox", "Panel", "GroupBox", "TabControl", "TabPage", "DataGridView",
        "ProgressBar", "TrackBar", "NumericUpDown", "DateTimePicker", "RichTextBox",
        "PictureBox", "MenuStrip", "ToolStrip", "StatusStrip", "SplitContainer",
        "FlowLayoutPanel", "TableLayoutPanel", "LinkLabel", "MaskedTextBox",
        "ListView", "WebBrowser", "MonthCalendar", "ContextMenuStrip",
        "Timer", "BindingSource", "DataSet", "ImageList", "ToolTip",
        "NotifyIcon", "ErrorProvider", "HelpProvider", "BackgroundWorker",
        "Form", "TreeView",
    ];
    for ct in &gui_ctors {
        let fn_name = format!("new_{}", ct);
        if let (Some(tid), Some(idx)) = (vm.type_registry.get_id(ct), h(vm, "vybe:gui", &fn_name)) {
            vm.type_registry.set_constructor(tid, Method::HostFn(idx));
        }
    }

    // ============================================================
    // Register enum types with compile-time constants
    // ============================================================
    register_enums(vm);

    // ============================================================
    // Export __tid_<name> globals for all registered types.
    // This allows compilers to emit set_type_id at construction sites.
    // ── New ECMA-262 + Web platform types ─────────────────────────
    register_new_globals_types(vm);

    // ============================================================
    for typedef in &vm.type_registry.types {
        let key = format!("__tid_{}", typedef.name.to_lowercase());
        if let Some(tid) = vm.type_registry.get_id(&typedef.name) {
            vm.globals.insert(key, Value::I32(tid as i32));
        }
    }
}

fn register_new_globals_types(vm: &mut VM) {
    let h = |vm: &VM, module: &str, name: &str| -> Option<usize> {
        vm.host_registry.get(&(module.to_string(), name.to_string())).copied()
    };

    // ── Iterator (Stage-3) ─────────────────────────────────────────
    {
        let mut t = TypeDef::new("Iterator");
        for (method, fname) in &[
            ("take", "take"), ("drop", "drop"), ("map", "map"), ("filter", "filter"),
            ("reduce", "reduce"), ("forEach", "forEach"), ("some", "some"),
            ("every", "every"), ("find", "find"), ("toArray", "toArray"),
            ("flatMap", "flatMap"),
        ] {
            if let Some(idx) = h(vm, "ecma:iterator", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // ── TextEncoder ────────────────────────────────────────────────
    {
        let mut t = TypeDef::new("TextEncoder");
        for (method, fname) in &[("encode", "encode"), ("encodeInto", "encodeInto")] {
            if let Some(idx) = h(vm, "web:encoding", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // ── TextDecoder ────────────────────────────────────────────────
    {
        let mut t = TypeDef::new("TextDecoder");
        if let Some(idx) = h(vm, "web:encoding", "decode") {
            t.methods.insert("decode".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // ── URLSearchParams ────────────────────────────────────────────
    {
        let mut t = TypeDef::new("URLSearchParams");
        for (method, fname) in &[
            ("get", "searchParamsGet"),
            ("has", "searchParamsHas"),
            ("toString", "searchParamsToString"),
        ] {
            if let Some(idx) = h(vm, "web:url", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // ── Response (fetch result) ────────────────────────────────────
    {
        let mut t = TypeDef::new("Response");
        for (method, fname) in &[("text", "responseText"), ("json", "responseJson")] {
            if let Some(idx) = h(vm, "web:fetch", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
}

fn register_enums(vm: &mut VM) {
    // DialogResult
    let id = vm.type_registry.register(TypeDef::new("DialogResult"));
    for (name, val) in &[("none",0),("ok",1),("cancel",2),("abort",3),("retry",4),("ignore",5),("yes",6),("no",7)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // MessageBoxButtons
    let id = vm.type_registry.register(TypeDef::new("MessageBoxButtons"));
    for (name, val) in &[("ok",0),("okcancel",1),("abortretryignore",2),("yesnocancel",3),("yesno",4),("retrycancel",5)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // MessageBoxIcon
    let id = vm.type_registry.register(TypeDef::new("MessageBoxIcon"));
    for (name, val) in &[("none",0),("error",16),("question",32),("warning",48),("information",64)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // Keys
    let id = vm.type_registry.register(TypeDef::new("Keys"));
    for (name, val) in &[
        ("none",0),("back",8),("tab",9),("return",13),("enter",13),("escape",27),
        ("space",32),("left",37),("up",38),("right",39),("down",40),
        ("delete",46),("insert",45),("shift",16),("control",17),("alt",18),
        ("f1",112),("f2",113),("f3",114),("f4",115),("f5",116),("f6",117),
        ("f7",118),("f8",119),("f9",120),("f10",121),("f11",122),("f12",123),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // FormBorderStyle
    let id = vm.type_registry.register(TypeDef::new("FormBorderStyle"));
    for (name, val) in &[("none",0),("fixedsingle",1),("fixeddialog",3),("sizable",4),("fixedtoolwindow",5),("sizabletoolwindow",6)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // FormStartPosition
    let id = vm.type_registry.register(TypeDef::new("FormStartPosition"));
    for (name, val) in &[("manual",0),("centerscreen",1),("windowsdefaultlocation",2),("windowsdefaultbounds",3),("centerparent",4)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // FormWindowState
    let id = vm.type_registry.register(TypeDef::new("FormWindowState"));
    for (name, val) in &[("normal",0),("minimized",1),("maximized",2)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // DockStyle
    let id = vm.type_registry.register(TypeDef::new("DockStyle"));
    for (name, val) in &[("none",0),("top",1),("bottom",2),("left",3),("right",4),("fill",5)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // AnchorStyles
    let id = vm.type_registry.register(TypeDef::new("AnchorStyles"));
    for (name, val) in &[("none",0),("top",1),("bottom",2),("left",4),("right",8)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // CloseReason
    let id = vm.type_registry.register(TypeDef::new("CloseReason"));
    for (name, val) in &[("none",0),("windowsshutdown",1),("userclosing",3),("applicationexitcall",5)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // MouseButtons
    let id = vm.type_registry.register(TypeDef::new("MouseButtons"));
    for (name, val) in &[("none",0),("left",1),("right",2),("middle",4)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // --- Exception types (cross-language: Python/JS/VB/C#/Dart) ---
    // Register base Exception type and common subtypes.
    // These enable `ref_test` and typed catch across all languages.
    let exc_base = vm.type_registry.register(TypeDef::new("Exception"));
    let exc_types = [
        "ValueError", "TypeError", "KeyError", "IndexError",
        "RuntimeError", "StopIteration", "AttributeError",
        "ZeroDivisionError", "FileNotFoundError", "ImportError",
        "NotImplementedError", "OverflowError", "IOError", "OSError",
        // .NET exception types
        "ArgumentException", "ArgumentNullException", "InvalidOperationException",
        "NullReferenceException", "FormatException", "StackOverflowException",
        // JS error types
        "Error", "RangeError", "SyntaxError", "ReferenceError", "URIError",
    ];
    for name in &exc_types {
        let mut td = TypeDef::new(name);
        td.parent = Some(exc_base);
        td.add_field("message");
        td.add_field("name");
        vm.type_registry.register(td);
    }
}
