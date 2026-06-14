//! Register built-in types in the VM's TypeRegistry.
//! Each type gets a vtable with methods resolved via the host function registry.
//! This replaces the legacy type_methods table with proper WASM GC-style dispatch.

use vybe_bytecode::{Method, TypeDef, VM, Value};

pub fn register_all(vm: &mut VM) {
    // Helper: look up host fn index by (module, name)
    let h = |vm: &VM, module: &str, name: &str| -> Option<usize> {
        vm.host_registry
            .get(&(module.to_string(), name.to_string()))
            .copied()
    };

    // --- Object (type 0, already created) ---
    // Universal methods bound to ECMA-262 primitives. `gethashcode` /
    // `equals` aren't in ECMA — using `String` coercion as a stand-in
    // until a typed equality / hash primitive lands.
    if let Some(idx) = h(vm, "ecma:string", "String") {
        vm.type_registry.add_host_method(0, "tostring", idx);
        vm.type_registry.add_host_method(0, "gethashcode", idx);
        vm.type_registry.add_host_method(0, "equals", idx);
    }
    // §20.1.3 Object.prototype methods — universal on every object so
    // `obj.hasOwnProperty` / `obj.isPrototypeOf` resolve as VALUES (not just
    // as direct calls). The lookup key is lowercased by the TypeRegistry.
    for (key, fname) in &[
        ("hasownproperty", "hasOwnProperty"),
        ("isprototypeof", "isPrototypeOf"),
        ("propertyisenumerable", "propertyIsEnumerable"),
        ("valueof", "valueOf"),
        ("tolocalestring", "toLocaleString"),
    ] {
        if let Some(idx) = h(vm, "ecma:object", fname) {
            vm.type_registry.add_host_method(0, key, idx);
        }
    }

    // --- String ---
    //
    // VB/C#/.NET-style instance method names (Contains, IndexOf,
    // Substring, etc.) bound directly to ECMA-262 §22.1
    // String.prototype. Same import surface JS engines satisfy
    // natively. Methods needing arg adaptation (.NET's `Insert`,
    // `Remove`, `LastIndexOf` with different-shape args) have
    // explicit adapter entries in emitter/dotnet/core/.
    let _string_id = {
        let mut t = TypeDef::new("String");
        for (method, module, fname) in &[
            ("contains", "ecma:string", "includes"),
            ("toupper", "ecma:string", "toUpperCase"),
            ("tolower", "ecma:string", "toLowerCase"),
            // JS-spec camelCase forms — TypeRegistry lowercases the
            // lookup key, so `s.toUpperCase()` and `s.toLowerCase()`
            // both arrive here as `touppercase` / `tolowercase`. These
            // entries make the JS-shape walker rewrites in php/walker.rs
            // (`ucfirst`, `lcfirst`, etc.) resolve through the same
            // ecma:string host fns the JS profile uses.
            ("touppercase", "ecma:string", "toUpperCase"),
            ("tolowercase", "ecma:string", "toLowerCase"),
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
            ("tostring", "ecma:string", "String"),
            ("toupperinvariant", "ecma:string", "toUpperCase"),
            ("tolowerinvariant", "ecma:string", "toLowerCase"),
            ("chars", "ecma:string", "charAt"),
            ("insert", "ecma:string", "substring"),
            ("remove", "ecma:string", "slice"),
            // ── JS-spec methods exposed under the same ecma:string surface
            // so `s.charAt(...)`, `s.charCodeAt(...)`, etc. dispatch via
            // the TypeRegistry without language-specific aliasing.
            ("charat", "ecma:string", "charAt"),
            ("charcodeat", "ecma:string", "charCodeAt"),
            ("codepointat", "ecma:string", "codePointAt"),
            ("at", "ecma:string", "at"),
            ("concat", "ecma:string", "concat"),
            ("includes", "ecma:string", "includes"),
            ("repeat", "ecma:string", "repeat"),
            ("padstart", "ecma:string", "padStart"),
            ("padend", "ecma:string", "padEnd"),
            ("normalize", "ecma:string", "normalize"),
            ("slice", "ecma:string", "slice"),
            ("substr", "ecma:string", "substr"),
            ("trimleft", "ecma:string", "trimStart"),
            ("trimright", "ecma:string", "trimEnd"),
            ("localecompare", "ecma:string", "localeCompare"),
            ("tolocalelowercase", "ecma:string", "toLowerCase"),
            ("tolocaleuppercase", "ecma:string", "toUpperCase"),
            // RegExp-driven String.prototype methods (§22.1.3.{11,13,14,17,18}).
            ("match", "ecma:regexp", "match"),
            ("matchall", "ecma:regexp", "matchAll"),
            ("search", "ecma:regexp", "search"),
            ("replaceall", "ecma:regexp", "replaceAll"),
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
    // the compile path (struct_get "count") auto-invokes ARRAY_LENGTH.
    // The method-call form `list.Count()` dispatches via TypeRegistry to
    // `ecma:array/length` which returns an i32.
    //
    // Constructor: `ecma:array.new` — plain JS Array (§23.1).
    // Range methods (InsertRange/RemoveRange/GetRange/SetRange/BinarySearch)
    // are compile-time adapters via `collections.*` common emitters.
    let list_id = {
        let mut t = TypeDef::new("List");
        if let Some(idx) = h(vm, "ecma:array", "new") {
            t.constructor = Some(Method::HostFn(idx));
        }
        for (method, module, fname) in &[
            ("add", "ecma:array", "push"),
            ("remove", "ecma:array", "removeValue"),
            ("removeat", "ecma:array", "removeAt"),
            ("contains", "ecma:array", "includes"),
            ("count", "ecma:array", "length"),
            ("clear", "ecma:array", "clear"),
            ("indexof", "ecma:array", "indexOf"),
            ("sort", "ecma:array", "sort"),
            ("reverse", "ecma:array", "reverse"),
            ("toarray", "ecma:array", "slice"),
            ("item", "ecma:array", "get"),
            ("lastindexof", "ecma:array", "lastIndexOf"),
            ("insert", "ecma:array", "insertAt"),
            ("addrange", "ecma:array", "concat"),
            ("capacity", "ecma:array", "length"),
            ("clone", "ecma:array", "slice"),
            ("copyto", "ecma:array", "slice"),
            ("trimtosize", "ecma:array", "length"),
            ("enqueue", "ecma:array", "push"),
            ("trydequeue", "ecma:array", "shift"),
            ("trypop", "ecma:array", "pop"),
            ("trypeek", "ecma:array", "last"),
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
    let _ = vm
        .type_registry
        .register(TypeDef::new("ArrayList").with_parent(list_id));
    let _ = vm
        .type_registry
        .register(TypeDef::new("Array").with_parent(list_id));

    // --- Dictionary (.NET adapter for ECMA-262 §24.1 Map) ---
    //
    // The Dictionary TypeDef is the .NET adapter layer. Its underlying
    // runtime shape is `ObjectKind::Map` (constructed by
    // `ecma:map.new`). Each .NET method name (Add / Item / ContainsKey
    // / Remove / Count / ...) is an alias that points at the
    // corresponding `ecma:map.*` host fn — the same primitive JS
    // `Map.prototype.*` calls. Adapter at the surface, ECMA underneath.
    let dict_id = {
        let mut t = TypeDef::new("Dictionary");
        for (method, fname) in &[
            ("add", "set"),                     // Dictionary.Add(k, v)
            ("item", "get"),                    // Dictionary.Item(k)
            ("containskey", "has"),             // Dictionary.ContainsKey(k)
            ("containsvalue", "containsValue"), // .NET-only linear scan
            ("remove", "delete"),               // Dictionary.Remove(k)
            ("keys", "keys"),
            ("values", "values"),
            ("clear", "clear"),
            ("count", "size"),      // .NET .Count maps to ECMA size
            ("trygetvalue", "get"), // 1-arg form (no `out` param)
            // ConcurrentDictionary aliases — same shape, same primitives.
            ("tryadd", "set"),
            ("addorupdate", "set"),
            ("getoradd", "get"),
        ] {
            if let Some(idx) = h(vm, "ecma:map", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t)
    };
    let _ = vm
        .type_registry
        .register(TypeDef::new("Hashtable").with_parent(dict_id));
    let _ = vm
        .type_registry
        .register(TypeDef::new("ConcurrentDictionary").with_parent(dict_id));
    let _ = vm
        .type_registry
        .register(TypeDef::new("SortedList").with_parent(dict_id));

    // --- Queue ---
    //
    // `.NET Queue<T>` is a JS Array used FIFO — `Enqueue` →
    // `push` appends at the end, `Dequeue` → `shift` removes from the
    // front. Property-style `q.Count` works via ARRAY_LENGTH.
    {
        let mut t = TypeDef::new("Queue");
        // .NET Queue<T> is a JS Array used FIFO. Methods route to
        // `ecma:array.*` directly; no `vybe:types` involvement.
        // `peek` looks at the front (`first`), `clear` empties.
        for (method, fname) in &[
            ("enqueue", "push"),
            ("dequeue", "shift"),
            ("peek", "first"), // FIFO: front
            ("count", "length"),
            ("clear", "clear"),
            ("contains", "includes"),
            ("toarray", "slice"),
        ] {
            if let Some(idx) = h(vm, "ecma:array", fname) {
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
        // .NET Stack<T> is a JS Array used LIFO. Methods route to
        // `ecma:array.*` directly; no `vybe:types` involvement.
        // `peek` looks at the end (`last`), `clear` empties.
        for (method, fname) in &[
            ("push", "push"),
            ("pop", "pop"),
            ("peek", "last"), // LIFO: top
            ("count", "length"),
            ("clear", "clear"),
            ("contains", "includes"),
            ("toarray", "slice"),
        ] {
            if let Some(idx) = h(vm, "ecma:array", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- HashSet ---
    //
    // Phase 7b: `.NET HashSet<T>` is a JS Set per ECMA-262 §24.2.
    // The mutating set-algebra methods (`UnionWith`, etc.) are .NET-shape
    // wrappers around the same `ObjectKind::Set` storage — distinct from
    // the immutable ES2025 `union`/`intersection`/etc. accessed via
    // `getMethodForCall` on JS Sets.
    {
        let mut t = TypeDef::new("HashSet");
        for (method, module, fname) in &[
            ("add", "ecma:set", "add"),
            ("contains", "ecma:set", "has"),
            ("remove", "ecma:set", "delete"),
            ("count", "ecma:set", "size"),
            ("clear", "ecma:set", "clear"),
            ("unionwith", "ecma:set", "unionWith"),
            ("intersectwith", "ecma:set", "intersectWith"),
            ("exceptwith", "ecma:set", "exceptWith"),
            ("symmetricexceptwith", "ecma:set", "symmetricExceptWith"),
            ("issubsetof", "ecma:set", "isSubsetOf"),
            ("issupersetof", "ecma:set", "isSupersetOf"),
            ("overlaps", "ecma:set", "overlaps"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // StringBuilder + DateTime TypeDef registrations retired — both
    // classes lower at compile time through the dotnet wrapper Component
    // Model adapters (`stringbuilder_adapter`, `datetime_adapter`). The
    // adapters compose `ecma:date.*`, `wasi:clocks/wall-clock.now`, and
    // inline `Op::DYN_ADD` string mutation directly on the wrapper objects.

    // --- SqlConnection ---
    {
        let mut t = TypeDef::new("SqlConnection");
        for (method, fname) in &[
            ("open", "[method]connection.open"),
            ("close", "[method]connection.close"),
            ("createcommand", "[method]connection.create-command"),
            ("begintransaction", "[method]connection.begin-transaction"),
            ("getschema", "[method]connection.get-schema"),
        ] {
            if let Some(idx) = h(vm, "wasi:sql/types", fname) {
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
            ("executenonquery", "[method]command.execute-non-query"),
            ("executescalar", "[method]command.execute-scalar"),
            ("executereader", "[method]command.execute-reader"),
            ("executenonqueryasync", "[method]command.execute-non-query"),
            ("executescalarasync", "[method]command.execute-scalar"),
            ("executereaderasync", "[method]command.execute-reader"),
            ("createparameter", "[method]command.create-parameter"),
        ] {
            if let Some(idx) = h(vm, "wasi:sql/types", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- SqlDataReader ---
    {
        let mut t = TypeDef::new("SqlDataReader");
        for (method, fname) in &[
            ("read", "[method]reader.read"),
            ("getvalue", "[method]reader.get-value"),
            ("getstring", "[method]reader.get-string"),
            ("getname", "[method]reader.get-name"),
            ("isdbnull", "[method]reader.is-dbnull"),
            ("close", "[method]reader.close"),
            ("dispose", "[method]reader.close"),
            ("getschematable", "[method]reader.get-schema-table"),
        ] {
            if let Some(idx) = h(vm, "wasi:sql/types", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- SqlTransaction ---
    {
        let mut t = TypeDef::new("SqlTransaction");
        for (method, fname) in &[
            ("commit", "[method]transaction.commit"),
            ("rollback", "[method]transaction.rollback"),
        ] {
            if let Some(idx) = h(vm, "wasi:sql/types", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- SqlDataAdapter ---
    {
        let mut t = TypeDef::new("SqlDataAdapter");
        if let Some(idx) = h(vm, "wasi:sql/types", "[method]adapter.fill") {
            t.methods.insert("fill".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- SqlParameterCollection ---
    {
        let mut t = TypeDef::new("SqlParameterCollection");
        for (method, fname) in &[
            ("addwithvalue", "[method]params.add-with-value"),
            ("clear", "[method]params.clear"),
            ("count", "[method]params.count"),
        ] {
            if let Some(idx) = h(vm, "wasi:sql/types", fname) {
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
            t.methods
                .insert("elapsedmilliseconds".into(), Method::HostFn(idx));
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
        if let Some(idx) = h(vm, "web:dom-parser", "serializeToString") {
            t.methods.insert("tostring".into(), Method::HostFn(idx));
            t.methods.insert("save".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- XElement ---
    {
        let mut t = TypeDef::new("XElement");
        if let Some(idx) = h(vm, "web:dom-parser", "serializeToString") {
            t.methods.insert("tostring".into(), Method::HostFn(idx));
            t.methods.insert("value".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- DataTable / DataSet / DataRow ---
    // Constructors and methods are compile-time adapters in
    // `emitter/dotnet/core/datatable_adapter.rs` — no runtime host fns.
    // TypeDefs kept for type-registry lookups only.
    {
        let mut t = TypeDef::new("DataTable");
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("DataSet");
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
    {
        let mut t = TypeDef::new("DataRow");
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
            ("set", "set"),
            ("get", "get"),
            ("has", "has"),
            ("delete", "delete"),
            ("keys", "keys"),
            ("values", "values"),
            ("clear", "clear"),
            ("entries", "entries"),
            ("forEach", "forEach"),
            ("iterator", "entries"),
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
            ("add", "add"),
            ("has", "has"),
            ("delete", "delete"),
            ("values", "values"),
            ("keys", "values"),
            ("clear", "clear"),
            ("entries", "entries"),
            ("forEach", "forEach"),
            ("iterator", "values"),
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
            ("set", "set"),
            ("get", "get"),
            ("has", "has"),
            ("delete", "delete"),
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
        for (method, fname) in &[("add", "add"), ("has", "has"), ("delete", "delete")] {
            if let Some(idx) = h(vm, "ecma:weakset", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- WeakRef (ECMA-262 §26.1) ---
    //
    // Bound to `ecma:weakref.*`. Strong-ref MVP stand-in (WASM GC MVP
    // doesn't expose weak refs); `deref()` always returns the target.
    {
        let mut t = TypeDef::new("WeakRef");
        for (method, fname) in &[("deref", "deref")] {
            if let Some(idx) = h(vm, "ecma:weakref", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- FinalizationRegistry (ECMA-262 §26.2) ---
    //
    // Bound to `ecma:finalizationregistry.*`. Strong-ref backing means
    // the cleanup callback never fires — API surface only.
    {
        let mut t = TypeDef::new("FinalizationRegistry");
        for (method, fname) in &[("register", "register"), ("unregister", "unregister")] {
            if let Some(idx) = h(vm, "ecma:finalizationregistry", fname) {
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
        for (method, fname) in &[("test", "test"), ("exec", "exec"), ("toString", "toString")] {
            if let Some(idx) = h(vm, "ecma:regexp", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Date (ECMA-262 §21.4) ---
    //
    // Backing object is a plain Object with `__time` (ms since epoch).
    // Methods all read/write that property via `ecma:date.*`. The TypeDef
    // entry here makes `d.getFullYear()` resolve through the vtable
    // instead of falling through to instance-property lookup (which would
    // miss because methods aren't stored on each Date instance).
    // `resolve_method` lowercases lookups, so insert keys lowercased.
    {
        let mut t = TypeDef::new("Date");
        for method in &[
            "getFullYear",
            "getYear",
            "getMonth",
            "getDate",
            "getDay",
            "getHours",
            "getMinutes",
            "getSeconds",
            "getMilliseconds",
            "getUTCFullYear",
            "getUTCMonth",
            "getUTCDate",
            "getUTCDay",
            "getUTCHours",
            "getUTCMinutes",
            "getUTCSeconds",
            "getUTCMilliseconds",
            "getTime",
            "getTimezoneOffset",
            "valueOf",
            "setTime",
            "setFullYear",
            "setMonth",
            "setDate",
            "setHours",
            "setMinutes",
            "setSeconds",
            "setMilliseconds",
            "setUTCFullYear",
            "setUTCMonth",
            "setUTCDate",
            "setUTCHours",
            "setUTCMinutes",
            "setUTCSeconds",
            "setUTCMilliseconds",
            "toISOString",
            "toString",
            "toUTCString",
            "toDateString",
            "toTimeString",
            "toJSON",
        ] {
            if let Some(idx) = h(vm, "ecma:date", method) {
                t.methods.insert(method.to_lowercase(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // PHP DateTime / DateTimeImmutable — re-use the ECMA Date method
    // surface so the walker's compile-time format pre-parse (which
    // emits `$dt.getFullYear()` / `$dt.getMonth()` / etc.) dispatches
    // through TypeRegistry to standard `ecma:date.*` host fns. No
    // PHP-specific method bindings — the adapter does any PHP-shaped
    // composition above the ECMA surface.
    for type_name in &["DateTime", "DateTimeImmutable"] {
        let mut t = TypeDef::new(type_name);
        for method in &[
            "getFullYear",
            "getMonth",
            "getDate",
            "getDay",
            "getHours",
            "getMinutes",
            "getSeconds",
            "getMilliseconds",
            "getTime",
            "getTimezoneOffset",
            "valueOf",
            "setTime",
            "setFullYear",
            "setMonth",
            "setDate",
            "setHours",
            "setMinutes",
            "setSeconds",
            "setMilliseconds",
            "toISOString",
            "toString",
        ] {
            if let Some(idx) = h(vm, "ecma:date", method) {
                t.methods.insert(method.to_lowercase(), Method::HostFn(idx));
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
            ("compare", "compare"),
            ("resolvedOptions", "resolvedOptions"),
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
            ("format", "format"),
            ("formatToParts", "formatToParts"),
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
            ("format", "format"),
            ("formatToParts", "formatToParts"),
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
            ("format", "format"),
            ("formatToParts", "formatToParts"),
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
            ("select", "select"),
            ("selectRange", "selectRange"),
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
            ("format", "format"),
            ("formatToParts", "formatToParts"),
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
            ("segment", "segment"),
            ("resolvedOptions", "resolvedOptions"),
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
            ("toString", "toString"),
            ("maximize", "maximize"),
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
        for (method, fname) in &[("of", "of"), ("resolvedOptions", "resolvedOptions")] {
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
            ("format", "format"),
            ("formatToParts", "formatToParts"),
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
    // .NET Process — methods (`Start`, `GetCurrentProcess`,
    // `WaitForExit`) all lower at compile time via Component Model
    // dispatch through `emitter::dotnet::core::process_adapter`. The
    // TypeDef is empty and serves only as a runtime hint for
    // `__type=Process` stamping.
    {
        let t = TypeDef::new("Process");
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
        "Button",
        "Label",
        "TextBox",
        "CheckBox",
        "RadioButton",
        "ComboBox",
        "ListBox",
        "Panel",
        "GroupBox",
        "TabControl",
        "TabPage",
        "DataGridView",
        "ProgressBar",
        "TrackBar",
        "NumericUpDown",
        "DateTimePicker",
        "RichTextBox",
        "PictureBox",
        "MenuStrip",
        "ToolStrip",
        "StatusStrip",
        "SplitContainer",
        "FlowLayoutPanel",
        "TableLayoutPanel",
        "LinkLabel",
        "MaskedTextBox",
        "ListView",
        "WebBrowser",
        "MonthCalendar",
        "ContextMenuStrip",
        "Timer",
        "BindingSource",
        "DataSet",
        "ImageList",
        "ToolTip",
        "NotifyIcon",
        "ErrorProvider",
        "HelpProvider",
        "BackgroundWorker",
        "TreeView",
    ];
    for ct in &control_type_names {
        vm.type_registry
            .register(TypeDef::new(ct).with_parent(control_id));
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
        // Runtime-hint constructor dispatch (compilation-hints proposal)
        // — fallback path when a type isn't statically known.
        //
        // Each .NET-shape collection points at its ECMA-262 primitive:
        //   Dictionary / Hashtable     → ECMA Map (§24.1)
        //   HashSet                     → ECMA Set (§24.2)
        //   Queue / Stack / List / etc. → ECMA Array (§23.1) — built
        //     by the compiler via `collections.new`/Op::ARRAY_NEW; no
        //     runtime ctor host fn (the bare `new Queue()` syntax in
        //     dynamic code paths is rare and routes through the dotnet
        //     wrapper Component Model dispatch at compile time).
        ("Dictionary", "ecma:map", "new"),
        ("HashSet", "ecma:set", "new"),
        ("SqlConnection", "wasi:sql/types", "connection.new"),
        ("SqlDataAdapter", "wasi:sql/types", "data-adapter.new"),
        // StringBuilder / DateTime / TcpClient / TcpListener / UdpClient /
        // Queue / Stack ctor mappings retired — all lower at compile time
        // through the dotnet wrapper Common emit path. Queue/Stack don't
        // need a runtime ctor since they're plain Arrays.
        ("Stopwatch", "wasi:clocks", "stopwatchNew"),
        // Random ctor: real WASI entropy. Seed argument (if any) is ignored —
        // wasi:random/insecure is a process-global PRNG with no per-instance
        // seed. The returned u64 becomes the "Random" receiver object; VB/C#
        // code treats it as opaque and calls methods on it. When compiler
        // lowering for range math lands, this becomes a trivial marker and
        // the method calls inline the WASI call themselves.
        ("Random", "wasi:random/insecure", "get-insecure-random-u64"),
        // DataTable / DataSet constructors lowered at compile time via
        // `emitter/dotnet/core/datatable_adapter.rs` — no runtime ctor host fn.
        ("Point", "vybe:gui", "pointNew"),
        ("Size", "vybe:gui", "sizeNew"),
        ("Font", "vybe:gui", "fontNew"),
    ];
    for (type_name, module, fname) in ctor_mappings {
        if let (Some(tid), Some(idx)) = (vm.type_registry.get_id(type_name), h(vm, module, fname)) {
            vm.type_registry.set_constructor(tid, Method::HostFn(idx));
        }
    }

    // Register GUI control constructors (new_Button, new_TextBox, etc.)
    let gui_ctors = [
        "Button",
        "Label",
        "TextBox",
        "CheckBox",
        "RadioButton",
        "ComboBox",
        "ListBox",
        "Panel",
        "GroupBox",
        "TabControl",
        "TabPage",
        "DataGridView",
        "ProgressBar",
        "TrackBar",
        "NumericUpDown",
        "DateTimePicker",
        "RichTextBox",
        "PictureBox",
        "MenuStrip",
        "ToolStrip",
        "StatusStrip",
        "SplitContainer",
        "FlowLayoutPanel",
        "TableLayoutPanel",
        "LinkLabel",
        "MaskedTextBox",
        "ListView",
        "WebBrowser",
        "MonthCalendar",
        "ContextMenuStrip",
        "Timer",
        "BindingSource",
        "DataSet",
        "ImageList",
        "ToolTip",
        "NotifyIcon",
        "ErrorProvider",
        "HelpProvider",
        "BackgroundWorker",
        "Form",
        "TreeView",
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
    // Register `__tid_<name>` globals for every built-in type. Built-ins
    // have canonical names like "String", "Object", "List", "Dictionary"
    // — we publish under BOTH the lowercased form (case-insensitive
    // languages: VB/Pascal/COBOL/PHP) AND the source-case form
    // (case-sensitive languages: JS/TS/Python/C#). Cross-language code
    // can resolve built-in types regardless of how the caller wrote them.
    for typedef in &vm.type_registry.types {
        if let Some(tid) = vm.type_registry.get_id(&typedef.name) {
            let lower = format!("__tid_{}", typedef.name.to_lowercase());
            vm.globals.insert(lower, Value::I32(tid as i32));
            let preserved = format!("__tid_{}", typedef.name);
            vm.globals
                .entry(preserved)
                .or_insert(Value::I32(tid as i32));
        }
    }
}

fn register_new_globals_types(vm: &mut VM) {
    let h = |vm: &VM, module: &str, name: &str| -> Option<usize> {
        vm.host_registry
            .get(&(module.to_string(), name.to_string()))
            .copied()
    };

    // ── Iterator (Stage-3) ─────────────────────────────────────────
    {
        let mut t = TypeDef::new("Iterator");
        for (method, fname) in &[
            ("take", "take"),
            ("drop", "drop"),
            ("map", "map"),
            ("filter", "filter"),
            ("reduce", "reduce"),
            ("forEach", "forEach"),
            ("some", "some"),
            ("every", "every"),
            ("find", "find"),
            ("toArray", "toArray"),
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
                t.methods.insert(method.to_lowercase(), Method::HostFn(idx));
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

    // --- WHATWG DOM types (web:dom-parser) -------------------------
    // Method tables for `Document` / `Element` so spec-shaped calls
    // (`elem.querySelector("...")`, `doc.getElementById("x")`,
    // `elem.getAttribute("href")`) dispatch through the TypeRegistry
    // vtable per Component Model resource semantics. Properties
    // (`tagName`, `nodeType`, `childNodes`, …) are set directly on the
    // node Object during parse and resolve via plain `Op::STRUCT_GET`
    // — only the *computed* methods need vtable entries here.
    // `TypeRegistry::resolve_method` lowercases the lookup key, so the
    // method-table keys must be lowercase too. Names like `querySelector`
    // here use the spec-cased form for documentation; the `to_lowercase`
    // call materialises the storage key.
    let element_id = {
        let mut t = TypeDef::new("Element");
        for (method, fname) in &[
            // Read API
            ("querySelector", "querySelector"),
            ("querySelectorAll", "querySelectorAll"),
            ("matches", "matches"),
            ("closest", "closest"),
            ("getAttribute", "getAttribute"),
            ("hasAttribute", "hasAttribute"),
            ("getElementsByTagName", "getElementsByTagName"),
            ("getElementsByClassName", "getElementsByClassName"),
            // Mutation API
            ("setAttribute", "setAttribute"),
            ("removeAttribute", "removeAttribute"),
            ("appendChild", "appendChild"),
            ("removeChild", "removeChild"),
            ("insertBefore", "insertBefore"),
            ("replaceChild", "replaceChild"),
            ("cloneNode", "cloneNode"),
            // Namespace-aware (Phase 4)
            ("getAttributeNS", "getAttributeNS"),
            ("hasAttributeNS", "hasAttributeNS"),
            ("setAttributeNS", "setAttributeNS"),
            ("removeAttributeNS", "removeAttributeNS"),
            ("getElementsByTagNameNS", "getElementsByTagNameNS"),
        ] {
            if let Some(idx) = h(vm, "web:dom-parser", fname) {
                t.methods.insert(method.to_lowercase(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t)
    };
    let document_id = {
        let mut t = TypeDef::new("Document");
        for (method, fname) in &[
            // Read API
            ("querySelector", "querySelector"),
            ("querySelectorAll", "querySelectorAll"),
            ("getElementById", "getElementById"),
            ("getElementsByTagName", "getElementsByTagName"),
            ("getElementsByClassName", "getElementsByClassName"),
            // Mutation factories
            ("createElement", "createElement"),
            ("createElementNS", "createElementNS"),
            ("createTextNode", "createTextNode"),
            ("createComment", "createComment"),
            ("createDocumentFragment", "createDocumentFragment"),
            ("appendChild", "appendChild"),
            ("removeChild", "removeChild"),
            ("insertBefore", "insertBefore"),
            ("replaceChild", "replaceChild"),
            ("cloneNode", "cloneNode"),
            ("getElementsByTagNameNS", "getElementsByTagNameNS"),
        ] {
            if let Some(idx) = h(vm, "web:dom-parser", fname) {
                t.methods.insert(method.to_lowercase(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t)
    };
    // Bare placeholder TypeDefs for the other DOM node kinds — no
    // method-table entries today (spec methods like
    // `Text.splitText(offset)` arrive in Phase 3 mutation work).
    let text_id = vm.type_registry.register(TypeDef::new("Text"));
    let comment_id = vm.type_registry.register(TypeDef::new("Comment"));
    let cdata_id = vm.type_registry.register(TypeDef::new("CDATASection"));
    let pi_id = vm
        .type_registry
        .register(TypeDef::new("ProcessingInstruction"));
    let attr_id = vm.type_registry.register(TypeDef::new("Attr"));
    let _ = vm.type_registry.register(TypeDef::new("DOMParser"));
    let _ = vm.type_registry.register(TypeDef::new("XMLSerializer"));
    let nnm_id = vm.type_registry.register(TypeDef::new("NamedNodeMap"));

    // Hand the IDs to `web::dom_parser` so the parser stamps each
    // constructed node's `Object::type_id` for vtable dispatch.
    crate::web::dom_parser::set_dom_type_ids(crate::web::dom_parser::DomTypeIds {
        document: document_id as usize,
        element: element_id as usize,
        text: text_id as usize,
        cdata: cdata_id as usize,
        comment: comment_id as usize,
        processing_instruction: pi_id as usize,
        attr: attr_id as usize,
        named_node_map: nnm_id as usize,
    });
}

fn register_enums(vm: &mut VM) {
    // DialogResult
    let id = vm.type_registry.register(TypeDef::new("DialogResult"));
    for (name, val) in &[
        ("none", 0),
        ("ok", 1),
        ("cancel", 2),
        ("abort", 3),
        ("retry", 4),
        ("ignore", 5),
        ("yes", 6),
        ("no", 7),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // MessageBoxButtons
    let id = vm.type_registry.register(TypeDef::new("MessageBoxButtons"));
    for (name, val) in &[
        ("ok", 0),
        ("okcancel", 1),
        ("abortretryignore", 2),
        ("yesnocancel", 3),
        ("yesno", 4),
        ("retrycancel", 5),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // MessageBoxIcon
    let id = vm.type_registry.register(TypeDef::new("MessageBoxIcon"));
    for (name, val) in &[
        ("none", 0),
        ("error", 16),
        ("question", 32),
        ("warning", 48),
        ("information", 64),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // Keys
    let id = vm.type_registry.register(TypeDef::new("Keys"));
    for (name, val) in &[
        ("none", 0),
        ("back", 8),
        ("tab", 9),
        ("return", 13),
        ("enter", 13),
        ("escape", 27),
        ("space", 32),
        ("left", 37),
        ("up", 38),
        ("right", 39),
        ("down", 40),
        ("delete", 46),
        ("insert", 45),
        ("shift", 16),
        ("control", 17),
        ("alt", 18),
        ("f1", 112),
        ("f2", 113),
        ("f3", 114),
        ("f4", 115),
        ("f5", 116),
        ("f6", 117),
        ("f7", 118),
        ("f8", 119),
        ("f9", 120),
        ("f10", 121),
        ("f11", 122),
        ("f12", 123),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // FormBorderStyle
    let id = vm.type_registry.register(TypeDef::new("FormBorderStyle"));
    for (name, val) in &[
        ("none", 0),
        ("fixedsingle", 1),
        ("fixeddialog", 3),
        ("sizable", 4),
        ("fixedtoolwindow", 5),
        ("sizabletoolwindow", 6),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // FormStartPosition
    let id = vm.type_registry.register(TypeDef::new("FormStartPosition"));
    for (name, val) in &[
        ("manual", 0),
        ("centerscreen", 1),
        ("windowsdefaultlocation", 2),
        ("windowsdefaultbounds", 3),
        ("centerparent", 4),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // FormWindowState
    let id = vm.type_registry.register(TypeDef::new("FormWindowState"));
    for (name, val) in &[("normal", 0), ("minimized", 1), ("maximized", 2)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // DockStyle
    let id = vm.type_registry.register(TypeDef::new("DockStyle"));
    for (name, val) in &[
        ("none", 0),
        ("top", 1),
        ("bottom", 2),
        ("left", 3),
        ("right", 4),
        ("fill", 5),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // AnchorStyles
    let id = vm.type_registry.register(TypeDef::new("AnchorStyles"));
    for (name, val) in &[
        ("none", 0),
        ("top", 1),
        ("bottom", 2),
        ("left", 4),
        ("right", 8),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // CloseReason
    let id = vm.type_registry.register(TypeDef::new("CloseReason"));
    for (name, val) in &[
        ("none", 0),
        ("windowsshutdown", 1),
        ("userclosing", 3),
        ("applicationexitcall", 5),
    ] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // MouseButtons
    let id = vm.type_registry.register(TypeDef::new("MouseButtons"));
    for (name, val) in &[("none", 0), ("left", 1), ("right", 2), ("middle", 4)] {
        vm.type_registry.add_constant(id, name, *val);
    }

    // --- Exception types (cross-language: Python/JS/VB/C#/Dart) ---
    // Register base Exception type and common subtypes.
    // These enable `ref_test` and typed catch across all languages.
    //
    // Hierarchy:
    //   Exception (root)
    //     Error (JS)
    //       TypeError, RangeError, SyntaxError, ReferenceError, URIError
    //     ValueError, KeyError, ... (Python/.NET — direct under Exception)
    let exc_base = vm.type_registry.register(TypeDef::new("Exception"));
    let mut error_td = TypeDef::new("Error");
    error_td.parent = Some(exc_base);
    error_td.add_field("message");
    error_td.add_field("name");
    let error_id = vm.type_registry.register(error_td);

    let exc_types_under_exception = [
        "ValueError",
        "KeyError",
        "IndexError",
        "RuntimeError",
        "StopIteration",
        "AttributeError",
        "ZeroDivisionError",
        "FileNotFoundError",
        "ImportError",
        "NotImplementedError",
        "OverflowError",
        "IOError",
        "OSError",
        // .NET exception types
        "ArgumentException",
        "ArgumentNullException",
        "InvalidOperationException",
        "NullReferenceException",
        "FormatException",
        "StackOverflowException",
    ];
    for name in &exc_types_under_exception {
        let mut td = TypeDef::new(name);
        td.parent = Some(exc_base);
        td.add_field("message");
        td.add_field("name");
        vm.type_registry.register(td);
    }

    // JS error subtypes — inherit from Error (per ECMA-262 §20.5.5).
    // TypeError is also used by Python for type errors; keep it under
    // Error so JS `instanceof Error` works while still being catchable
    // as Exception (Error's parent) in cross-language code.
    let js_error_types = [
        "TypeError",
        "RangeError",
        "SyntaxError",
        "ReferenceError",
        "URIError",
    ];
    for name in &js_error_types {
        let mut td = TypeDef::new(name);
        td.parent = Some(error_id);
        td.add_field("message");
        td.add_field("name");
        vm.type_registry.register(td);
    }
}
