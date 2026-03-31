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
    let _string_id = {
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
        if let Some(idx) = h(vm, "vybe:types", "listNew") {
            t.constructor = Some(Method::HostFn(idx));
        }
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
            ("lastindexof", "vybe:types", "listIndexOf"),
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
    {
        let mut t = TypeDef::new("Map");
        for (method, fname) in &[
            ("set", "mapSet"), ("get", "mapGet"), ("has", "mapHas"),
            ("delete", "mapDelete"), ("keys", "mapKeys"), ("clear", "mapClear"),
        ] {
            if let Some(idx) = h(vm, "vybe:collections", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        // size/count via listCount
        if let Some(idx) = h(vm, "vybe:types", "listCount") {
            t.methods.insert("size".to_string(), Method::HostFn(idx));
            t.methods.insert("count".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        vm.type_registry.register(t);
    }

    // --- Set (JS collection) ---
    {
        let mut t = TypeDef::new("Set");
        for (method, fname) in &[
            ("add", "setAdd"), ("has", "setHas"), ("delete", "setDelete"),
            ("values", "setValues"), ("clear", "setClear"),
        ] {
            if let Some(idx) = h(vm, "vybe:collections", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        if let Some(idx) = h(vm, "vybe:types", "listCount") {
            t.methods.insert("size".to_string(), Method::HostFn(idx));
            t.methods.insert("count".to_string(), Method::HostFn(idx));
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
    {
        let t = TypeDef::new("Task");
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
        ("TcpClient", "vybe:net", "tcpConnect"),
        ("TcpListener", "vybe:net", "tcpListenerNew"),
        ("UdpClient", "vybe:net", "udpNew"),
        ("StreamReader", "vybe:net", "streamReaderNew"),
        ("StreamWriter", "vybe:net", "streamWriterNew"),
        ("Stopwatch", "vybe:threading", "stopwatchNew"),
        ("Random", "vybe:threading", "randomNew"),
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
    // ============================================================
    for typedef in &vm.type_registry.types {
        let key = format!("__tid_{}", typedef.name.to_lowercase());
        if let Some(tid) = vm.type_registry.get_id(&typedef.name) {
            vm.globals.insert(key, Value::I32(tid as i32));
        }
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
