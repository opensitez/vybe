//! Register built-in types in the VM's TypeRegistry.
//! Each type gets a vtable with methods resolved via the host function registry.
//! This replaces the legacy type_methods table with proper WASM GC-style dispatch.

use vybe_bytecode::{VM, TypeDef, Method};

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
    let string_id = {
        let mut t = TypeDef::new("String");
        for (method, module, fname) in &[
            ("contains", "vybe:string", "includes"),
            ("toupper", "vybe:string", "toUpperCase"),
            ("tolower", "vybe:string", "toLowerCase"),
            ("trim", "vybe:string", "trim"),
            ("trimstart", "vybe:string", "trimStart"),
            ("trimend", "vybe:string", "trimEnd"),
            ("startswith", "vybe:string", "startsWith"),
            ("endswith", "vybe:string", "endsWith"),
            ("indexof", "vybe:string", "indexOf"),
            ("lastindexof", "vybe:string", "indexOf"),
            ("substring", "vybe:string", "substring"),
            ("replace", "vybe:string", "replace"),
            ("split", "vybe:string", "split"),
            ("padleft", "vybe:string", "padStart"),
            ("padright", "vybe:string", "padEnd"),
            ("tostring", "vybe:convert", "toString"),
            ("toupperinvariant", "vybe:string", "toUpperCase"),
            ("tolowerinvariant", "vybe:string", "toLowerCase"),
            ("chars", "vybe:string", "charAt"),
            ("insert", "vybe:string", "mid"),
            ("remove", "vybe:string", "slice"),
        ] {
            if let Some(idx) = h(vm, module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0); // inherits from Object
        vm.type_registry.register(t)
    };

    // --- List / Array ---
    let list_id = {
        let mut t = TypeDef::new("List");
        for (method, module, fname) in &[
            ("add", "vybe:types", "listAdd"),
            ("remove", "vybe:types", "listRemove"),
            ("removeat", "vybe:types", "listRemoveAt"),
            ("contains", "vybe:types", "listContains"),
            ("count", "vybe:types", "listCount"),
            ("clear", "vybe:types", "listClear"),
            ("indexof", "vybe:types", "listIndexOf"),
            ("sort", "vybe:types", "listSort"),
            ("reverse", "vybe:types", "listReverse"),
            ("toarray", "vybe:types", "listToArray"),
            ("item", "vybe:types", "listItem"),
            ("insert", "vybe:types", "listInsert"),
            ("addrange", "vybe:types", "listAddRange"),
            ("capacity", "vybe:types", "listCount"),
            // Array methods
            ("push", "vybe:array", "push"),
            ("pop", "vybe:array", "pop"),
            ("shift", "vybe:array", "shift"),
            ("join", "vybe:array", "join"),
            ("includes", "vybe:array", "includes"),
            ("slice", "vybe:array", "slice"),
            ("concat", "vybe:array", "concat"),
            ("find", "vybe:types", "listIndexOf"),
            ("findindex", "vybe:types", "listIndexOf"),
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
    let dict_id = {
        let mut t = TypeDef::new("Dictionary");
        for (method, module, fname) in &[
            ("add", "vybe:types", "dictAdd"),
            ("item", "vybe:types", "dictItem"),
            ("containskey", "vybe:types", "dictContainsKey"),
            ("containsvalue", "vybe:types", "dictContainsKey"),
            ("remove", "vybe:types", "dictRemove"),
            ("keys", "vybe:types", "dictKeys"),
            ("values", "vybe:types", "dictValues"),
            ("clear", "vybe:types", "dictClear"),
            ("trygetvalue", "vybe:types", "dictItem"),
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
    {
        let mut t = TypeDef::new("Queue");
        for (method, fname) in &[
            ("enqueue", "queueEnqueue"), ("dequeue", "queueDequeue"), ("peek", "queuePeek"),
            ("count", "listCount"), ("clear", "listClear"), ("contains", "listContains"),
            ("toarray", "listToArray"),
        ] {
            if let Some(idx) = h(vm, "vybe:types", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Stack ---
    {
        let mut t = TypeDef::new("Stack");
        for (method, fname) in &[
            ("push", "stackPush"), ("pop", "stackPop"), ("peek", "stackPeek"),
            ("count", "listCount"), ("clear", "listClear"), ("contains", "listContains"),
        ] {
            if let Some(idx) = h(vm, "vybe:types", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- HashSet ---
    {
        let mut t = TypeDef::new("HashSet");
        for (method, fname) in &[
            ("add", "hashSetAdd"), ("contains", "hashSetContains"), ("remove", "hashSetRemove"),
            ("count", "listCount"), ("clear", "listClear"),
        ] {
            if let Some(idx) = h(vm, "vybe:types", fname) {
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
            ("close", "tcpClose"), ("getstream", "tcpGetStream"),
        ] {
            if let Some(idx) = h(vm, "vybe:net", fname) {
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
            ("accepttcpclient", "tcpListenerAccept"),
        ] {
            if let Some(idx) = h(vm, "vybe:net", fname) {
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
            ("send", "udpSend"), ("receive", "udpReceive"), ("close", "udpClose"),
        ] {
            if let Some(idx) = h(vm, "vybe:net", fname) {
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
            ("readline", "streamReaderReadLine"), ("readtoend", "streamReaderReadLine"),
            ("close", "streamWriterClose"),
        ] {
            if let Some(idx) = h(vm, "vybe:net", fname) {
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
            ("writeline", "streamWriterWriteLine"), ("write", "streamWriterWriteLine"),
            ("flush", "streamWriterFlush"), ("close", "streamWriterClose"),
        ] {
            if let Some(idx) = h(vm, "vybe:net", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Stopwatch ---
    {
        let mut t = TypeDef::new("Stopwatch");
        if let Some(idx) = h(vm, "vybe:threading", "stopwatchElapsed") {
            t.methods.insert("elapsedmilliseconds".into(), Method::HostFn(idx));
            t.methods.insert("elapsed".into(), Method::HostFn(idx));
            t.methods.insert("stop".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Random ---
    {
        let mut t = TypeDef::new("Random");
        if let Some(idx) = h(vm, "vybe:threading", "randomNext") {
            t.methods.insert("next".into(), Method::HostFn(idx));
        }
        if let Some(idx) = h(vm, "vybe:threading", "randomNextDouble") {
            t.methods.insert("nextdouble".into(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }
}
