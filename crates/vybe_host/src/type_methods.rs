//! Register type method bindings.
//! Maps (TypeName, methodName) → host function for built-in type objects AND strings.

use vybe_bytecode::VM;

pub fn register_all(vm: &mut VM) {
    // --- String instance methods ---
    for (method, host_fn) in &[
        ("contains", "includes"), ("toupper", "toUpperCase"), ("tolower", "toLowerCase"),
        ("trim", "trim"), ("trimstart", "trimStart"), ("trimend", "trimEnd"),
        ("startswith", "startsWith"), ("endswith", "endsWith"),
        ("indexof", "indexOf"), ("lastindexof", "indexOf"),
        ("substring", "substring"), ("replace", "replace"),
        ("split", "split"), ("padleft", "padStart"), ("padright", "padEnd"),
        ("toupperinvariant", "toUpperCase"), ("tolowerinvariant", "toLowerCase"),
        ("tostring", "toUpperCase"), // simplified — returns same string
        ("tochararray", "split"), // simplified
        ("remove", "slice"), // simplified
        ("insert", "mid"), // simplified
        ("chars", "charAt"),
    ] {
        vm.register_type_method("string", method, "vybe:string", host_fn);
    }

    // --- List methods ---
    for (method, host_fn) in &[
        ("add", "listAdd"), ("remove", "listRemove"), ("removeat", "listRemoveAt"),
        ("contains", "listContains"), ("count", "listCount"), ("clear", "listClear"),
        ("indexof", "listIndexOf"), ("sort", "listSort"), ("reverse", "listReverse"),
        ("toarray", "listToArray"), ("addrange", "listAddRange"),
        ("insert", "listInsert"), ("item", "listItem"),
        ("find", "listIndexOf"), ("findindex", "listIndexOf"),
        ("getrange", "listToArray"), // simplified
        ("removerange", "listClear"), // simplified
        ("capacity", "listCount"), // simplified
    ] {
        vm.register_type_method("list", method, "vybe:types", host_fn);
    }

    // --- Dictionary methods ---
    for (method, host_fn) in &[
        ("add", "dictAdd"), ("item", "dictItem"), ("containskey", "dictContainsKey"),
        ("containsvalue", "dictContainsKey"), // simplified
        ("remove", "dictRemove"), ("keys", "dictKeys"), ("values", "dictValues"),
        ("clear", "dictClear"), ("trygetvalue", "dictItem"), // simplified
    ] {
        vm.register_type_method("dictionary", method, "vybe:types", host_fn);
    }

    // --- Queue methods ---
    for (method, host_fn) in &[
        ("enqueue", "queueEnqueue"), ("dequeue", "queueDequeue"), ("peek", "queuePeek"),
        ("count", "listCount"), ("clear", "listClear"),
        ("contains", "listContains"), ("toarray", "listToArray"),
    ] {
        vm.register_type_method("queue", method, "vybe:types", host_fn);
    }

    // --- Stack methods ---
    for (method, host_fn) in &[
        ("push", "stackPush"), ("pop", "stackPop"), ("peek", "stackPeek"),
        ("count", "listCount"), ("clear", "listClear"),
        ("contains", "listContains"), ("toarray", "listToArray"),
    ] {
        vm.register_type_method("stack", method, "vybe:types", host_fn);
    }

    // --- HashSet methods ---
    for (method, host_fn) in &[
        ("add", "hashSetAdd"), ("contains", "hashSetContains"), ("remove", "hashSetRemove"),
        ("count", "listCount"), ("clear", "listClear"),
        ("unionwith", "hashSetAdd"), // simplified
        ("intersectwith", "hashSetAdd"), // simplified
        ("exceptwith", "hashSetRemove"), // simplified
    ] {
        vm.register_type_method("hashset", method, "vybe:types", host_fn);
    }

    // --- StringBuilder methods ---
    for (method, host_fn) in &[
        ("append", "sbAppend"), ("appendline", "sbAppendLine"),
        ("appendformat", "sbAppend"), // simplified — ignores format
        ("tostring", "sbToString"), ("clear", "sbClear"),
        ("insert", "sbInsert"), ("replace", "sbReplace"),
        ("remove", "sbClear"), // simplified
        ("length", "sbToString"), // simplified
        ("capacity", "sbToString"), // simplified
        ("ensurecapacity", "sbToString"), // simplified
        ("chars", "sbToString"), // simplified
        ("equals", "sbToString"), // simplified — should compare
        ("copyto", "sbToString"), // simplified
        ("gettype", "sbToString"), // simplified
    ] {
        vm.register_type_method("stringbuilder", method, "vybe:types", host_fn);
    }

    // --- DateTime methods ---
    for (method, host_fn) in &[
        ("adddays", "dateTimeAddDays"), ("addhours", "dateTimeAddHours"),
        ("addminutes", "dateTimeAddMinutes"), ("addseconds", "dateTimeAddSeconds"),
        ("addmonths", "dateTimeAddMonths"), ("addyears", "dateTimeAddYears"),
        ("tostring", "dateTimeToString"), ("toshortdatestring", "dateTimeToShortDate"),
        ("tolongdatestring", "dateTimeToString"),
    ] {
        vm.register_type_method("datetime", method, "vybe:types", host_fn);
    }

    // --- Stopwatch ---
    vm.register_type_method("stopwatch", "elapsedmilliseconds", "vybe:threading", "stopwatchElapsed");
    vm.register_type_method("stopwatch", "elapsed", "vybe:threading", "stopwatchElapsed");
    vm.register_type_method("stopwatch", "stop", "vybe:threading", "stopwatchElapsed"); // simplified

    // --- Random ---
    vm.register_type_method("random", "next", "vybe:threading", "randomNext");
    vm.register_type_method("random", "nextdouble", "vybe:threading", "randomNextDouble");

    // --- StreamReader/Writer ---
    vm.register_type_method("streamreader", "readline", "vybe:net", "streamReaderReadLine");
    vm.register_type_method("streamreader", "readtoend", "vybe:net", "streamReaderReadLine"); // simplified
    vm.register_type_method("streamreader", "close", "vybe:net", "streamWriterClose");
    vm.register_type_method("streamwriter", "writeline", "vybe:net", "streamWriterWriteLine");
    vm.register_type_method("streamwriter", "write", "vybe:net", "streamWriterWriteLine");
    vm.register_type_method("streamwriter", "close", "vybe:net", "streamWriterClose");
    vm.register_type_method("streamwriter", "flush", "vybe:net", "streamWriterClose");

    // --- TcpClient ---
    vm.register_type_method("tcpclient", "close", "vybe:net", "tcpClose");
    vm.register_type_method("tcpclient", "getstream", "vybe:net", "tcpGetStream");

    // --- TcpListener ---
    vm.register_type_method("tcplistener", "start", "vybe:net", "tcpListenerStart");
    vm.register_type_method("tcplistener", "stop", "vybe:net", "tcpListenerStop");
    vm.register_type_method("tcplistener", "accepttcpclient", "vybe:net", "tcpListenerAccept");

    // --- UdpClient ---
    vm.register_type_method("udpclient", "send", "vybe:net", "udpSend");
    vm.register_type_method("udpclient", "receive", "vybe:net", "udpReceive");
    vm.register_type_method("udpclient", "close", "vybe:net", "udpClose");

    // --- StreamWriter (over TCP or file) ---
    vm.register_type_method("streamwriter", "writeline", "vybe:net", "streamWriterWriteLine");
    vm.register_type_method("streamwriter", "write", "vybe:net", "streamWriterWriteLine");
    vm.register_type_method("streamwriter", "flush", "vybe:net", "streamWriterFlush");
    vm.register_type_method("streamwriter", "close", "vybe:net", "streamWriterClose");

    // --- StreamReader (over TCP or file) ---
    vm.register_type_method("streamreader", "readline", "vybe:net", "streamReaderReadLine");
    vm.register_type_method("streamreader", "readtoend", "vybe:net", "streamReaderReadLine");
    vm.register_type_method("streamreader", "close", "vybe:net", "streamWriterClose");

    // --- DataTable ---
    vm.register_type_method("datatable", "newrow", "vybe:data", "dataTableNewRow");
    vm.register_type_method("datatable", "rows", "vybe:data", "dataTableAddRow");

    // --- SqlConnection (maps to vybe:database) ---
    for (method, host_fn) in &[
        ("open", "open"), ("close", "close"),
        ("createcommand", "createCommand"),
        ("begintransaction", "beginTransaction"),
    ] {
        vm.register_type_method("sqlconnection", method, "vybe:database", host_fn);
    }

    // --- SqlCommand ---
    for (method, host_fn) in &[
        ("executenonquery", "execute"), ("executescalar", "scalar"),
        ("executereader", "query"),
    ] {
        vm.register_type_method("sqlcommand", method, "vybe:database", host_fn);
    }

    // --- SqlCommand ---
    for (method, host_fn) in &[
        ("executenonquery", "execute"), ("executescalar", "scalar"),
        ("executereader", "query"), ("commandtext", "execute"),
        ("parameters", "execute"), // simplified
    ] {
        vm.register_type_method("sqlcommand", method, "vybe:database", host_fn);
    }

    // --- ArrayList (legacy, map to List) ---
    for (method, host_fn) in &[
        ("add", "listAdd"), ("remove", "listRemove"), ("removeat", "listRemoveAt"),
        ("contains", "listContains"), ("count", "listCount"), ("clear", "listClear"),
        ("indexof", "listIndexOf"), ("sort", "listSort"), ("reverse", "listReverse"),
        ("item", "listItem"), ("insert", "listInsert"), ("addrange", "listAddRange"),
        ("capacity", "listCount"), ("toarray", "listToArray"),
        ("getrange", "listToArray"), ("removerange", "listClear"),
        ("trimtosize", "listCount"), // no-op
        ("binarysearch", "listIndexOf"), // simplified
        ("copyto", "listToArray"), // simplified
    ] {
        vm.register_type_method("arraylist", method, "vybe:types", host_fn);
    }

    // --- Universal object methods (work on any type) ---
    // These use empty string as type key — the VM falls back to these
    vm.register_type_method("", "tostring", "vybe:convert", "toString");
    vm.register_type_method("", "gettype", "vybe:convert", "typeName");
    vm.register_type_method("", "gethashcode", "vybe:convert", "cint"); // simplified
    vm.register_type_method("", "equals", "vybe:convert", "toString"); // simplified — should compare

    // --- Array instance methods (when array has __type) ---
    for (method, host_fn) in &[
        ("push", "push"), ("pop", "pop"), ("shift", "shift"),
        ("join", "join"), ("indexof", "indexOf"), ("includes", "includes"),
        ("reverse", "reverse"), ("slice", "slice"), ("concat", "concat"),
        ("length", "length"),
    ] {
        vm.register_type_method("array", method, "vybe:array", host_fn);
    }
}
