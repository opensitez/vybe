//! Register type method bindings.
//! Maps (TypeName, methodName) → host function for built-in type objects.
//! This allows obj.Method() to work when the method isn't a property on the object.

use vybe_bytecode::VM;

pub fn register_all(vm: &mut VM) {
    // List methods
    for (method, host_fn) in &[
        ("add", "listAdd"), ("remove", "listRemove"), ("removeat", "listRemoveAt"),
        ("contains", "listContains"), ("count", "listCount"), ("clear", "listClear"),
        ("indexof", "listIndexOf"), ("sort", "listSort"), ("reverse", "listReverse"),
        ("toarray", "listToArray"),
    ] {
        vm.register_type_method("list", method, "vybe:types", host_fn);
    }

    // Dictionary methods
    for (method, host_fn) in &[
        ("add", "dictAdd"), ("item", "dictItem"), ("containskey", "dictContainsKey"),
        ("remove", "dictRemove"), ("keys", "dictKeys"), ("values", "dictValues"),
        ("clear", "dictClear"),
    ] {
        vm.register_type_method("dictionary", method, "vybe:types", host_fn);
    }

    // Queue methods
    for (method, host_fn) in &[
        ("enqueue", "queueEnqueue"), ("dequeue", "queueDequeue"), ("peek", "queuePeek"),
    ] {
        vm.register_type_method("queue", method, "vybe:types", host_fn);
    }

    // Stack methods
    for (method, host_fn) in &[
        ("push", "stackPush"), ("pop", "stackPop"), ("peek", "stackPeek"),
    ] {
        vm.register_type_method("stack", method, "vybe:types", host_fn);
    }

    // HashSet methods
    for (method, host_fn) in &[
        ("add", "hashSetAdd"), ("contains", "hashSetContains"), ("remove", "hashSetRemove"),
    ] {
        vm.register_type_method("hashset", method, "vybe:types", host_fn);
    }

    // StringBuilder methods
    for (method, host_fn) in &[
        ("append", "sbAppend"), ("appendline", "sbAppendLine"),
        ("tostring", "sbToString"), ("clear", "sbClear"),
        ("insert", "sbInsert"), ("replace", "sbReplace"),
    ] {
        vm.register_type_method("stringbuilder", method, "vybe:types", host_fn);
    }

    // DateTime methods
    for (method, host_fn) in &[
        ("adddays", "dateTimeAddDays"), ("addhours", "dateTimeAddHours"),
        ("addminutes", "dateTimeAddMinutes"), ("addseconds", "dateTimeAddSeconds"),
        ("addmonths", "dateTimeAddMonths"), ("addyears", "dateTimeAddYears"),
        ("tostring", "dateTimeToString"), ("toshortdatestring", "dateTimeToShortDate"),
    ] {
        vm.register_type_method("datetime", method, "vybe:types", host_fn);
    }

    // Stopwatch methods
    vm.register_type_method("stopwatch", "elapsedmilliseconds", "vybe:threading", "stopwatchElapsed");
    vm.register_type_method("stopwatch", "elapsed", "vybe:threading", "stopwatchElapsed");

    // Random methods
    vm.register_type_method("random", "next", "vybe:threading", "randomNext");
    vm.register_type_method("random", "nextdouble", "vybe:threading", "randomNextDouble");

    // StreamReader methods
    vm.register_type_method("streamreader", "readline", "vybe:net", "streamReaderReadLine");

    // StreamWriter methods
    vm.register_type_method("streamwriter", "writeline", "vybe:net", "streamWriterWriteLine");
    vm.register_type_method("streamwriter", "close", "vybe:net", "streamWriterClose");

    // TcpClient methods
    vm.register_type_method("tcpclient", "send", "vybe:net", "tcpSend");
    vm.register_type_method("tcpclient", "receive", "vybe:net", "tcpReceive");
    vm.register_type_method("tcpclient", "close", "vybe:net", "tcpClose");
}
