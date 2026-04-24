use std::sync::LazyLock;

use super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef};

pub fn class_exports() -> &'static [DotnetClassExport] {
    static EXPORTS: LazyLock<Vec<DotnetClassExport>> = LazyLock::new(|| {
        vec![
            // Phase 7b: List<T>'s runtime method dispatch goes through
            // `ecma:array/*` (see vybe_host/src/builtin_types.rs).
            // The constructor stays on `vybe:types/listNew` because
            // that's where `__get_count` / `__get_length` auto-getters
            // get installed so VB `list.Count` (property-style read)
            // auto-invokes the getter. Once a typed `Common` ctor path
            // that can install auto-getters lands, the ctor can move
            // to `collections.new`.
            collection_class(
                "dotnet.System.Collections.Generic",
                "List",
                "vybe:types",
                "listNew",
                &[
                    ("Add", 1, "vybe:types", "listAdd"),
                    ("Remove", 1, "vybe:types", "listRemove"),
                    ("RemoveAt", 1, "vybe:types", "listRemoveAt"),
                    ("Contains", 1, "vybe:types", "listContains"),
                    ("Count", 0, "vybe:types", "listCount"),
                    ("Clear", 0, "vybe:types", "listClear"),
                    ("IndexOf", 1, "vybe:types", "listIndexOf"),
                    ("Sort", 0, "vybe:types", "listSort"),
                    ("Reverse", 0, "vybe:types", "listReverse"),
                    ("ToArray", 0, "vybe:types", "listToArray"),
                    ("Item", 1, "vybe:types", "listItem"),
                    ("Insert", 2, "vybe:types", "listInsert"),
                    ("AddRange", 1, "vybe:types", "listAddRange"),
                ],
            ),
            collection_class(
                "dotnet.System.Collections.Generic",
                "Dictionary",
                "vybe:types",
                "dictNew",
                &[
                    ("Add", 2, "vybe:types", "dictAdd"),
                    ("Item", 1, "vybe:types", "dictItem"),
                    ("ContainsKey", 1, "vybe:types", "dictContainsKey"),
                    ("Remove", 1, "vybe:types", "dictRemove"),
                    ("Keys", 0, "vybe:types", "dictKeys"),
                    ("Values", 0, "vybe:types", "dictValues"),
                    ("Clear", 0, "vybe:types", "dictClear"),
                ],
            ),
            constructor_class("dotnet.System.Collections.Generic", "Queue", "vybe:types", "queueNew"),
            constructor_class("dotnet.System.Collections.Generic", "Stack", "vybe:types", "stackNew"),
            constructor_class("dotnet.System.Collections.Generic", "HashSet", "vybe:types", "hashSetNew"),
            DotnetClassExport::new(
                "dotnet.System.Collections.Concurrent",
                ClassType::new("ConcurrentDictionary")
                    .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("vybe:types", "dictNew")))
                    .with_method(MethodDef::new("TryAdd",      2, MethodBody::Common("dict.set_dynamic".into())))
                    .with_method(MethodDef::new("TryGetValue", 2, MethodBody::Common("dict.get_dynamic".into())))
                    .with_method(MethodDef::new("AddOrUpdate", 3, MethodBody::Common("dict.set_dynamic".into())))
                    .with_method(MethodDef::new("GetOrAdd",    2, MethodBody::Common("dict.get_dynamic".into())))
                    .with_method(MethodDef::new("ContainsKey", 1, MethodBody::Common("dict.has".into())))
                    .with_method(MethodDef::new("Remove",      1, MethodBody::Common("dict.delete".into())))
                    .with_method(MethodDef::new("Clear",       0, MethodBody::Common("dict.clear".into())))
                    .with_method(MethodDef::new("Count",       0, MethodBody::Common("dict.size".into()))),
            ),
            DotnetClassExport::new(
                "dotnet.System.Collections.Concurrent",
                ClassType::new("ConcurrentQueue")
                    .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("vybe:types", "listNew")))
                    .with_method(MethodDef::new("Enqueue",    1, MethodBody::Common("collections.push".into())))
                    .with_method(MethodDef::new("TryDequeue", 1, MethodBody::Common("collections.shift".into())))
                    .with_method(MethodDef::new("TryPeek",    1, MethodBody::Common("collections.get".into())))
                    .with_method(MethodDef::new("Count",      0, MethodBody::Common("collections.length".into()))),
            ),
            DotnetClassExport::new(
                "dotnet.System.Collections.Concurrent",
                ClassType::new("ConcurrentStack")
                    .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("vybe:types", "listNew")))
                    .with_method(MethodDef::new("Push",    1, MethodBody::Common("collections.push".into())))
                    .with_method(MethodDef::new("TryPop",  0, MethodBody::Common("collections.pop".into())))
                    .with_method(MethodDef::new("TryPeek", 0, MethodBody::Common("collections.get".into())))
                    .with_method(MethodDef::new("Count",   0, MethodBody::Common("collections.length".into()))),
            ),
            constructor_class("dotnet.System.Collections.Generic", "SortedList", "vybe:types", "dictNew"),
            common_constructor_class("dotnet.System.Collections.Generic", "LinkedList", "collections.new"),
            DotnetClassExport::new(
                "dotnet.System.Collections",
                ClassType::new("ArrayList")
                    .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                    .with_method(MethodDef::new("Add",           1, MethodBody::Common("collections.push".into())))
                    .with_method(MethodDef::new("Remove",        1, MethodBody::Common("collections.remove".into())))
                    .with_method(MethodDef::new("RemoveAt",      1, MethodBody::Common("collections.remove_at".into())))
                    .with_method(MethodDef::new("Contains",      1, MethodBody::Common("collections.contains".into())))
                    .with_method(MethodDef::new("Count",         0, MethodBody::Common("collections.length".into())))
                    .with_method(MethodDef::new("Capacity",      0, MethodBody::Common("collections.length".into())))
                    .with_method(MethodDef::new("Clear",         0, MethodBody::Common("collections.clear".into())))
                    .with_method(MethodDef::new("IndexOf",       1, MethodBody::Common("collections.index_of".into())))
                    .with_method(MethodDef::new("IndexOf2",      2, MethodBody::Common("collections.index_of_from".into())))
                    .with_method(MethodDef::new("LastIndexOf",   1, MethodBody::Common("collections.last_index_of".into())))
                    .with_method(MethodDef::new("LastIndexOf2",  2, MethodBody::Common("collections.last_index_of_from".into())))
                    .with_method(MethodDef::new("Sort",          0, MethodBody::Common("collections.sort".into())))
                    .with_method(MethodDef::new("Reverse",       0, MethodBody::Common("collections.reverse".into())))
                    .with_method(MethodDef::new("ReverseRange",  2, MethodBody::Common("collections.reverse_range".into())))
                    .with_method(MethodDef::new("ToArray",       0, MethodBody::Common("collections.clone".into())))
                    .with_method(MethodDef::new("Clone",         0, MethodBody::Common("collections.clone".into())))
                    .with_method(MethodDef::new("Item",          1, MethodBody::Common("collections.get".into())))
                    .with_method(MethodDef::new("Insert",        2, MethodBody::Common("collections.insert".into())))
                    .with_method(MethodDef::new("InsertRange",   2, MethodBody::Common("collections.insert_range".into())))
                    .with_method(MethodDef::new("RemoveRange",   2, MethodBody::Common("collections.remove_range".into())))
                    .with_method(MethodDef::new("GetRange",      2, MethodBody::Common("collections.get_range".into())))
                    .with_method(MethodDef::new("SetRange",      2, MethodBody::Common("collections.set_range".into())))
                    .with_method(MethodDef::new("BinarySearch",  1, MethodBody::Common("collections.binary_search".into())))
                    .with_method(MethodDef::new("AddRange",      1, MethodBody::Common("collections.concat".into()))),
            ),
            constructor_class("dotnet.System.Collections", "Hashtable", "vybe:types", "dictNew"),
            constructor_class("dotnet.System.Collections", "Collection", "vybe:types", "listNew"),
            constructor_class("dotnet.System", "DateTime", "vybe:types", "dateTimeNew"),
            constructor_class("dotnet.System", "Random", "wasi:random/insecure", "get-insecure-random-u64"),
            static_only_class(
                "dotnet.System",
                "Console",
                &[
                    ("WriteLine", 1, "wasi:cli", "log"),
                    ("Write", 1, "wasi:cli", "log"),
                    ("ReadLine", 0, "wasi:cli", "readLine"),
                    ("Error", 1, "wasi:cli", "error"),
                    ("Print", 1, "wasi:cli", "log"),
                    ("Assert", 1, "wasi:cli", "log"),
                ],
            ),
            static_only_class(
                "dotnet.System",
                "Convert",
                &[
                    ("ToInt32", 1, "vybe:convert", "cint"),
                    ("ToDouble", 1, "vybe:convert", "cdbl"),
                    ("ToString", 1, "vybe:convert", "toString"),
                    ("ToBoolean", 1, "vybe:convert", "cbool"),
                    ("ToDateTime", 1, "vybe:convert", "toString"),
                ],
            ),
            static_only_class(
                "dotnet.System",
                "String",
                &[("Format", 2, "vybe:string", "format")],
            ),
            static_only_class(
                "dotnet.System",
                "Environment",
                &[("GetEnvironmentVariable", 1, "wasi:cli", "getEnv")],
            ),
            DotnetClassExport::new(
                "dotnet.System.Threading.Tasks",
                ClassType::new("Task")
                    .with_method(MethodDef::static_method(
                        "Run",
                        1,
                        MethodBody::Common("threading.task_run".to_string()),
                    ))
                    .with_method(MethodDef::static_method(
                        "Delay",
                        1,
                        // Task.Delay(ms) → vybe:runtime/taskDelay. Returns a
                        // real async Task whose `iscompleted` flag transitions
                        // false → true after `ms` milliseconds via a native
                        // thread-spawn on the host. The WASI-aligned primitives
                        // underneath are wasi:clocks/monotonic-clock (duration
                        // timing) and a wasi.thread-spawn-shaped background
                        // worker (implemented with std::thread in-host).
                        MethodBody::HostCall(HostTarget::new("vybe:runtime", "taskDelay")),
                    )),
            ),
            DotnetClassExport::new(
                "dotnet.System.Threading",
                ClassType::new("Interlocked")
                    .with_method(MethodDef::static_method(
                        "Add",
                        2,
                        MethodBody::Common("threading.atomic_add".to_string()),
                    ))
                    .with_method(MethodDef::static_method(
                        "Exchange",
                        2,
                        MethodBody::Common("threading.atomic_xchg".to_string()),
                    ))
                    .with_method(MethodDef::static_method(
                        "CompareExchange",
                        3,
                        MethodBody::Common("threading.atomic_cmpxchg".to_string()),
                    )),
            ),
            constructor_class("dotnet.System.Text", "StringBuilder", "vybe:types", "stringBuilderNew"),
            constructor_class("dotnet.System.Data", "DataTable", "vybe:data", "dataTableNew"),
            constructor_class("dotnet.System.Data", "DataSet", "vybe:data", "dataSetNew"),
            constructor_class("dotnet.System.Drawing", "Point", "vybe:drawing", "pointNew"),
            constructor_class("dotnet.System.Drawing", "Size", "vybe:drawing", "sizeNew"),
            constructor_class("dotnet.System.Drawing", "SizeF", "vybe:drawing", "sizeNew"),
            constructor_class("dotnet.System.Drawing", "Font", "vybe:drawing", "fontNew"),
            constructor_class("dotnet.System.Drawing", "Pen", "vybe:drawing", "penNew"),
            constructor_class("dotnet.System.Drawing", "SolidBrush", "vybe:drawing", "solidBrushNew"),
            constructor_class("dotnet.System.Drawing", "Color", "vybe:drawing", "colorFromName"),
            constructor_class("dotnet.System.Drawing", "Graphics", "vybe:drawing", "graphicsNew"),
            constructor_and_static_class(
                "dotnet.System.Diagnostics",
                "Stopwatch",
                Some(("wasi:clocks", "stopwatchNew")),
                &[("StartNew", 0, "wasi:clocks", "stopwatchNew")],
            ),
            static_only_class(
                "dotnet.System.Diagnostics",
                "Debug",
                &[
                    ("WriteLine", 1, "wasi:cli", "log"),
                    ("Write", 1, "wasi:cli", "log"),
                    ("Assert", 1, "wasi:cli", "log"),
                ],
            ),
            static_only_class(
                "dotnet.System.Diagnostics",
                "Trace",
                &[("WriteLine", 1, "wasi:cli", "log")],
            ),
            constructor_class("dotnet.System.Diagnostics", "ProcessStartInfo", "vybe:types", "processStartInfoNew"),
            constructor_and_static_class(
                "dotnet.System.Diagnostics",
                "Process",
                Some(("vybe:types", "processNew")),
                &[
                    ("Start", 1, "vybe:types", "processStart"),
                    ("GetCurrentProcess", 0, "vybe:types", "processGetCurrent"),
                ],
            ),
            constructor_class("dotnet.System.Data.SqlClient", "SqlConnection", "vybe:database", "connect"),
            constructor_class("dotnet.System.Net.Sockets", "TcpClient", "dotnet:sockets", "tcpClientNew"),
            constructor_class("dotnet.System.Net.Sockets", "TcpListener", "dotnet:sockets", "tcpListenerNew"),
            constructor_class("dotnet.System.Net.Sockets", "UdpClient", "dotnet:sockets", "udpClientNew"),
            static_only_class(
                "dotnet.System.Net",
                "Dns",
                &[
                    ("GetHostAddresses", 1, "dotnet:net", "dnsGetHostAddresses"),
                    ("GetHostEntry", 1, "dotnet:net", "dnsGetHostEntry"),
                    ("GetHostName", 0, "dotnet:net", "dnsGetHostName"),
                ],
            ),
            collection_class(
                "dotnet.System.IO",
                "StreamReader",
                "dotnet:io",
                "streamReaderNew",
                &[
                    ("ReadLine",   0, "dotnet:io", "streamReaderReadLine"),
                    ("ReadToEnd",  0, "dotnet:io", "streamReaderReadToEnd"),
                    ("Close",      0, "dotnet:io", "streamWriterClose"),
                    ("Dispose",    0, "dotnet:io", "streamWriterClose"),
                ],
            ),
            collection_class(
                "dotnet.System.IO",
                "StreamWriter",
                "dotnet:io",
                "streamWriterNew",
                &[
                    ("WriteLine",  1, "dotnet:io", "streamWriterWriteLine"),
                    ("Write",      1, "dotnet:io", "streamWriterWrite"),
                    ("Flush",      0, "dotnet:io", "streamWriterFlush"),
                    ("Close",      0, "dotnet:io", "streamWriterClose"),
                    ("Dispose",    0, "dotnet:io", "streamWriterClose"),
                ],
            ),
            static_only_class(
                "dotnet.System.IO",
                "File",
                &[
                    ("ReadAllText", 1, "wasi:filesystem", "readFile"),
                    ("WriteAllText", 2, "wasi:filesystem", "writeFile"),
                    ("AppendAllText", 2, "wasi:filesystem", "appendFile"),
                    ("Exists", 1, "wasi:filesystem", "exists"),
                    ("Delete", 1, "wasi:filesystem", "remove"),
                    ("Copy", 2, "wasi:filesystem", "copy"),
                    ("Move", 2, "wasi:filesystem", "rename"),
                ],
            ),
            static_only_class(
                "dotnet.System.IO",
                "Directory",
                &[
                    ("CreateDirectory", 1, "wasi:filesystem", "mkdir"),
                    ("GetFiles", 1, "wasi:filesystem", "listDir"),
                ],
            ),
            static_only_class(
                "dotnet.System.IO",
                "Path",
                &[
                    ("Combine", 2, "wasi:filesystem", "pathCombine"),
                    ("GetFileName", 1, "wasi:filesystem", "pathGetFileName"),
                    ("GetExtension", 1, "wasi:filesystem", "pathGetExtension"),
                    ("GetDirectoryName", 1, "wasi:filesystem", "pathGetDirectory"),
                    ("GetFileNameWithoutExtension", 1, "wasi:filesystem", "pathGetFileNameWithoutExt"),
                    ("ChangeExtension", 2, "wasi:filesystem", "pathChangeExtension"),
                    ("GetFullPath", 1, "wasi:filesystem", "pathGetFullPath"),
                    ("GetTempPath", 0, "wasi:filesystem", "pathGetTempPath"),
                ],
            ),
            DotnetClassExport::new(
                "dotnet.System.Threading",
                ClassType::new("Thread")
                    .with_method(MethodDef::static_method(
                        "Sleep",
                        1,
                        MethodBody::Common("threading.sleep".to_string()),
                    )),
            ),
        ]
    });
    EXPORTS.as_slice()
}

pub fn component_class_exports() -> &'static [(&'static str, ClassType)] {
    static EXPORTS: LazyLock<Vec<(&'static str, ClassType)>> = LazyLock::new(|| {
        class_exports()
            .iter()
            .map(|export| (export.interface, export.class.clone()))
            .collect()
    });
    EXPORTS.as_slice()
}

fn constructor_class(
    interface: &'static str,
    name: &'static str,
    module: &'static str,
    ctor: &'static str,
) -> DotnetClassExport {
    DotnetClassExport::new(
        interface,
        ClassType::new(name)
            .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new(module, ctor))),
    )
}

fn common_constructor_class(
    interface: &'static str,
    name: &'static str,
    emit: &'static str,
) -> DotnetClassExport {
    DotnetClassExport::new(
        interface,
        ClassType::new(name)
            .with_constructor(ConstructorDef::new(0).with_common_backing(emit)),
    )
}

fn constructor_and_static_class(
    interface: &'static str,
    name: &'static str,
    constructor: Option<(&'static str, &'static str)>,
    methods: &[(&'static str, u8, &'static str, &'static str)],
) -> DotnetClassExport {
    let mut class = ClassType::new(name);
    if let Some((module, ctor)) = constructor {
        class = class.with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new(module, ctor)));
    }
    for (method, arity, module, func) in methods {
        class = class.with_method(MethodDef::static_method(
            *method,
            *arity,
            MethodBody::HostCall(HostTarget::new(*module, *func)),
        ));
    }
    DotnetClassExport::new(interface, class)
}

fn collection_class(
    interface: &'static str,
    name: &'static str,
    module: &'static str,
    ctor: &'static str,
    methods: &[(&'static str, u8, &'static str, &'static str)],
) -> DotnetClassExport {
    let mut class = ClassType::new(name)
        .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new(module, ctor)));
    for (method, arity, method_module, method_name) in methods {
        class = class.with_method(MethodDef::new(
            *method,
            *arity,
            MethodBody::HostCall(HostTarget::new(*method_module, *method_name)),
        ));
    }
    DotnetClassExport::new(interface, class)
}

/// Like `collection_class` but every method binds to a `compiler_common`
/// primitive (`Common("…")`) instead of a host function. Used for
/// `.NET` classes whose semantics map 1:1 to the shared JS-shape
/// primitives (Array, Map, Set) — `List<T>` → `collections.*`,
/// `Dictionary<K,V>` → `dict.*`, etc. The provider of the primitive
/// is one-file-swappable per the `feedback_compiler_common_is_THE_emitter`
/// rule; rewiring `collections.push` from `ecma:array.push` to a
/// polyfill happens in `compiler_common/collections.rs` alone.
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

fn static_only_class(
    interface: &'static str,
    name: &'static str,
    methods: &[(&'static str, u8, &'static str, &'static str)],
) -> DotnetClassExport {
    constructor_and_static_class(interface, name, None, methods)
}