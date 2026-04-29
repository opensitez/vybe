use std::sync::LazyLock;

use super::super::class_exports::DotnetClassExport;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef};

pub fn class_exports() -> &'static [DotnetClassExport] {
    static EXPORTS: LazyLock<Vec<DotnetClassExport>> = LazyLock::new(|| {
        vec![
            // .NET `List<T>` is shape-identical to ECMA-262 §23.1 Array.
            // The constructor materializes a real `ObjectKind::Array` via
            // `collections.new` (Op::ARRAY_NEW) and every method routes
            // through the corresponding `collections.*` primitive (which
            // itself routes to `ecma:array.*` per the WASM spec). The
            // .NET-name → ECMA-name translation is the wrapper's job.
            collection_class_common(
                "dotnet.System.Collections.Generic",
                "List",
                "collections.new",
                &[
                    ("Add",         1, "collections.push"),
                    ("Remove",      1, "collections.remove"),
                    ("RemoveAt",    1, "collections.remove_at"),
                    ("Contains",    1, "collections.contains"),
                    ("Count",       0, "collections.length"),
                    ("Clear",       0, "collections.clear"),
                    ("IndexOf",     1, "collections.index_of"),
                    ("Sort",        0, "collections.sort"),
                    ("Reverse",     0, "collections.reverse"),
                    ("ToArray",     0, "collections.clone"),
                    ("Item",        1, "collections.get"),
                    ("Insert",      2, "collections.insert"),
                    ("AddRange",    1, "collections.concat"),
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
                    .with_method(MethodDef::new("Add",         2, MethodBody::HostCall(HostTarget::new("ecma:map", "set"))))
                    .with_method(MethodDef::new("Item",        1, MethodBody::HostCall(HostTarget::new("ecma:map", "get"))))
                    .with_method(MethodDef::new("ContainsKey", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "has"))))
                    .with_method(MethodDef::new("Remove",      1, MethodBody::HostCall(HostTarget::new("ecma:map", "delete"))))
                    .with_method(MethodDef::new("Keys",        0, MethodBody::HostCall(HostTarget::new("ecma:map", "keys"))))
                    .with_method(MethodDef::new("Values",      0, MethodBody::HostCall(HostTarget::new("ecma:map", "values"))))
                    .with_method(MethodDef::new("Clear",       0, MethodBody::HostCall(HostTarget::new("ecma:map", "clear"))))
                    .with_method(MethodDef::new("Count",       0, MethodBody::HostCall(HostTarget::new("ecma:map", "size")))),
            ),
            // .NET `Queue<T>` is a JS Array used FIFO — `Enqueue` appends
            // (push), `Dequeue` removes from the front (shift), `Peek`
            // looks at the front (`ecma:array.first`).
            collection_class_common(
                "dotnet.System.Collections.Generic",
                "Queue",
                "collections.new",
                &[
                    ("Enqueue",  1, "collections.push"),
                    ("Dequeue",  0, "collections.shift"),
                    ("Count",    0, "collections.length"),
                    ("Clear",    0, "collections.clear"),
                    ("Contains", 1, "collections.contains"),
                    ("ToArray",  0, "collections.clone"),
                ],
            ),
            // .NET `Stack<T>` is a JS Array used LIFO — `Push` appends
            // (push), `Pop` removes from the end (pop), `Peek` looks at
            // the end (`ecma:array.last`).
            collection_class_common(
                "dotnet.System.Collections.Generic",
                "Stack",
                "collections.new",
                &[
                    ("Push",     1, "collections.push"),
                    ("Pop",      0, "collections.pop"),
                    ("Count",    0, "collections.length"),
                    ("Clear",    0, "collections.clear"),
                    ("Contains", 1, "collections.contains"),
                    ("ToArray",  0, "collections.clone"),
                ],
            ),
            // .NET `HashSet<T>` is a real ECMA-262 §24.2 `Set`. Constructor
            // creates an `ObjectKind::Set`; methods route through the
            // matching `ecma:set.*` host fns.
            DotnetClassExport::new(
                "dotnet.System.Collections.Generic",
                ClassType::new("HashSet")
                    .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:set", "new")))
                    .with_method(MethodDef::new("Add",      1, MethodBody::HostCall(HostTarget::new("ecma:set", "add"))))
                    .with_method(MethodDef::new("Remove",   1, MethodBody::HostCall(HostTarget::new("ecma:set", "delete"))))
                    .with_method(MethodDef::new("Contains", 1, MethodBody::HostCall(HostTarget::new("ecma:set", "has"))))
                    .with_method(MethodDef::new("Count",    0, MethodBody::HostCall(HostTarget::new("ecma:set", "size"))))
                    .with_method(MethodDef::new("Clear",    0, MethodBody::HostCall(HostTarget::new("ecma:set", "clear")))),
            ),
            // `ConcurrentDictionary` is a thread-safe `Dictionary` — same
            // shape (ECMA Map). Atomicity isn't modeled; methods route the
            // same way.
            DotnetClassExport::new(
                "dotnet.System.Collections.Concurrent",
                ClassType::new("ConcurrentDictionary")
                    .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:map", "new")))
                    .with_method(MethodDef::new("TryAdd",      2, MethodBody::HostCall(HostTarget::new("ecma:map", "set"))))
                    .with_method(MethodDef::new("TryGetValue", 2, MethodBody::HostCall(HostTarget::new("ecma:map", "get"))))
                    .with_method(MethodDef::new("AddOrUpdate", 3, MethodBody::HostCall(HostTarget::new("ecma:map", "set"))))
                    .with_method(MethodDef::new("GetOrAdd",    2, MethodBody::HostCall(HostTarget::new("ecma:map", "get"))))
                    .with_method(MethodDef::new("ContainsKey", 1, MethodBody::HostCall(HostTarget::new("ecma:map", "has"))))
                    .with_method(MethodDef::new("Remove",      1, MethodBody::HostCall(HostTarget::new("ecma:map", "delete"))))
                    .with_method(MethodDef::new("Clear",       0, MethodBody::HostCall(HostTarget::new("ecma:map", "clear"))))
                    .with_method(MethodDef::new("Count",       0, MethodBody::HostCall(HostTarget::new("ecma:map", "size")))),
            ),
            // ConcurrentQueue / ConcurrentStack — same shape as their
            // non-concurrent counterparts (Array). Atomicity isn't
            // modeled at this layer.
            DotnetClassExport::new(
                "dotnet.System.Collections.Concurrent",
                ClassType::new("ConcurrentQueue")
                    .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                    .with_method(MethodDef::new("Enqueue",    1, MethodBody::Common("collections.push".into())))
                    .with_method(MethodDef::new("TryDequeue", 1, MethodBody::Common("collections.shift".into())))
                    .with_method(MethodDef::new("TryPeek",    1, MethodBody::Common("collections.get".into())))
                    .with_method(MethodDef::new("Count",      0, MethodBody::Common("collections.length".into()))),
            ),
            DotnetClassExport::new(
                "dotnet.System.Collections.Concurrent",
                ClassType::new("ConcurrentStack")
                    .with_constructor(ConstructorDef::new(0).with_common_backing("collections.new"))
                    .with_method(MethodDef::new("Push",    1, MethodBody::Common("collections.push".into())))
                    .with_method(MethodDef::new("TryPop",  0, MethodBody::Common("collections.pop".into())))
                    .with_method(MethodDef::new("TryPeek", 0, MethodBody::Common("collections.get".into())))
                    .with_method(MethodDef::new("Count",   0, MethodBody::Common("collections.length".into()))),
            ),
            constructor_class("dotnet.System.Collections.Generic", "SortedList", "ecma:map", "new"),
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
            constructor_class("dotnet.System.Collections", "Hashtable", "ecma:map", "new"),
            common_constructor_class("dotnet.System.Collections", "Collection", "collections.new"),
            // .NET `DateTime` is shape-identical to ECMA-262 §21.4 `Date`.
            // Constructor materializes a real Date via `ecma:date/new`,
            // which itself reads through `wasi:clocks/wall-clock.now`
            // for parameterless construction. Static methods (Now /
            // UtcNow / Today / Parse) lower through
            // `emitter::dotnet::core::datetime_adapter` —
            // `ecma:date.now` / `ecma:date.parse` (which read
            // `wasi:clocks/wall-clock.now`) wrapped in a `{__type:
            // "DateTime", __time: ms}` object.
            DotnetClassExport::new(
                "dotnet.System",
                ClassType::new("DateTime")
                    .with_constructor(ConstructorDef::new(0).with_backing(HostTarget::new("ecma:date", "new")))
                    .with_method(MethodDef::static_method("Now",     0, MethodBody::Common("dotnet.datetime_now".into())))
                    .with_method(MethodDef::static_method("UtcNow",  0, MethodBody::Common("dotnet.datetime_now".into())))
                    .with_method(MethodDef::static_method("Today",   0, MethodBody::Common("dotnet.datetime_today".into())))
                    .with_method(MethodDef::static_method("Parse",   1, MethodBody::Common("dotnet.datetime_parse".into()))),
            ),
            constructor_class("dotnet.System", "Random", "wasi:random/insecure", "get-insecure-random-u64"),
            // .NET `System.TimeSpan` factory statics — unit-to-ms
            // conversion + record build. Pure inline bytecode via
            // `emitter::dotnet::core::timespan_adapter`. No host fns.
            DotnetClassExport::new(
                "dotnet.System",
                ClassType::new("TimeSpan")
                    .with_method(MethodDef::static_method("FromDays",         1, MethodBody::Common("dotnet.timespan_from_days".into())))
                    .with_method(MethodDef::static_method("FromHours",        1, MethodBody::Common("dotnet.timespan_from_hours".into())))
                    .with_method(MethodDef::static_method("FromMinutes",      1, MethodBody::Common("dotnet.timespan_from_minutes".into())))
                    .with_method(MethodDef::static_method("FromSeconds",      1, MethodBody::Common("dotnet.timespan_from_seconds".into())))
                    .with_method(MethodDef::static_method("FromMilliseconds", 1, MethodBody::Common("dotnet.timespan_from_milliseconds".into())))
                    .with_method(MethodDef::static_method("Zero",             0, MethodBody::Common("dotnet.timespan_zero".into()))),
            ),
            // .NET `System.Array` static methods — range operations
            // (Clear/Copy/Resize) and Sort that don't have 1:1 ECMA
            // mirrors. Each lowers through
            // `emitter::dotnet::core::array_adapter` to a stdlib
            // bytecode chunk composing `ecma:array.*` primitives.
            // `Reverse` / `IndexOf` route to ecma:array directly.
            DotnetClassExport::new(
                "dotnet.System",
                ClassType::new("Array")
                    .with_method(MethodDef::static_method("Clear",   3, MethodBody::Common("dotnet.array_clear".into())))
                    .with_method(MethodDef::static_method("Copy",    3, MethodBody::Common("dotnet.array_copy".into())))
                    .with_method(MethodDef::static_method("Resize",  2, MethodBody::Common("dotnet.array_resize".into())))
                    .with_method(MethodDef::static_method("Sort",    1, MethodBody::Common("dotnet.array_sort".into())))
                    .with_method(MethodDef::static_method("Reverse", 1, MethodBody::HostCall(HostTarget::new("ecma:array", "reverse"))))
                    .with_method(MethodDef::static_method("IndexOf", 2, MethodBody::HostCall(HostTarget::new("ecma:array", "indexOf")))),
            ),
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
                    ("ToDouble", 1, "ecma:number", "Number"),
                    ("ToString", 1, "ecma:string", "String"),
                    ("ToBoolean", 1, "ecma:boolean", "Boolean"),
                    ("ToDateTime", 1, "ecma:string", "String"),
                ],
            ),
            // .NET `System.String` static methods — `Format` lowers
            // through `emitter::dotnet::core::string_format_adapter` to
            // inline bytecode. No host fns. The dispatch threads `argc`
            // so any number of placeholder args works.
            DotnetClassExport::new(
                "dotnet.System",
                ClassType::new("String")
                    .with_method(MethodDef::static_method(
                        "Format",
                        2,
                        MethodBody::Common("dotnet.string_format".into()),
                    )),
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
                        // Task.Delay(ms) lowers to `Op::THREAD_SPAWN` of a
                        // tiny worker that calls `wasi:clocks/sleep(ms)`.
                        // The Task object — `iscompleted` / `isalive` /
                        // `result` / `status` — is constructed by the VM's
                        // native THREAD_SPAWN handler when the worker
                        // returns. Pure WASM: wasi-threads `thread.spawn`
                        // opcode + WASI `wasi:clocks/sleep` import — zero
                        // host fns.
                        MethodBody::Common("threading.task_delay".into()),
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
            // .NET StringBuilder — no direct ECMA mirror, but the wrapper
            // materializes a plain Object with a `__buffer` string and
            // mutates via `Op::DYN_ADD` (string concat) + `STRUCT_SET`.
            // Pure WASM, zero host fns. The constructor is overloaded
            // (`StringBuilder()` and `StringBuilder("initial")`); the
            // dispatch layer threads `argc` so the single Common emit
            // picks the right shape at compile time.
            DotnetClassExport::new(
                "dotnet.System.Text",
                ClassType::new("StringBuilder")
                    .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.string_builder_new"))
                    .with_method(MethodDef::new("Append",     1, MethodBody::Common("dotnet.sb_append".into())))
                    .with_method(MethodDef::new("AppendLine", 1, MethodBody::Common("dotnet.sb_append_line".into())))
                    .with_method(MethodDef::new("ToString",   0, MethodBody::Common("dotnet.sb_to_string".into())))
                    .with_method(MethodDef::new("Clear",      0, MethodBody::Common("dotnet.sb_clear".into())))
                    .with_method(MethodDef::new("Length",     0, MethodBody::Common("dotnet.sb_length".into())))
                    .with_method(MethodDef::new("Insert",     2, MethodBody::Common("dotnet.sb_insert".into())))
                    .with_method(MethodDef::new("Replace",    2, MethodBody::Common("dotnet.sb_replace".into()))),
            ),
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
            // .NET Process / ProcessStartInfo — lowers to
            // `node:child_process.spawnSync` (Node-shape child process
            // primitive) plus plain Object structs for the .NET-shape
            // ProcessStartInfo record. See
            // `emitter::dotnet::core::process_adapter`. Multi-arity
            // ProcessStartInfo ctors handled via threaded `argc`.
            common_constructor_class(
                "dotnet.System.Diagnostics",
                "ProcessStartInfo",
                "dotnet.process_start_info_new",
            ),
            DotnetClassExport::new(
                "dotnet.System.Diagnostics",
                ClassType::new("Process")
                    .with_constructor(ConstructorDef::new(0).with_common_backing("dotnet.process_new"))
                    .with_method(MethodDef::static_method("Start",             1, MethodBody::Common("dotnet.process_start".into())))
                    .with_method(MethodDef::static_method("GetCurrentProcess", 0, MethodBody::Common("dotnet.process_get_current".into())))
                    .with_method(MethodDef::new("WaitForExit",                 0, MethodBody::Common("dotnet.process_wait_for_exit".into()))),
            ),
            constructor_class("dotnet.System.Data.SqlClient", "SqlConnection", "vybe:database", "connect"),
            // .NET sockets — adapter at the wrapper layer routes every
            // method through `wasi:sockets/*` (TCP/UDP/IP name lookup)
            // + `wasi:io/streams.*` for the stream-shaped methods. No
            // `dotnet:*` host module involved. See
            // `emitter::dotnet::core::sockets_adapter`.
            DotnetClassExport::new(
                "dotnet.System.Net.Sockets",
                ClassType::new("TcpClient")
                    .with_constructor(ConstructorDef::new(2).with_common_backing("dotnet.tcp_client_new"))
                    .with_method(MethodDef::new("GetStream", 0, MethodBody::Common("dotnet.tcp_client_get_stream".into())))
                    .with_method(MethodDef::new("Close",     0, MethodBody::Common("dotnet.tcp_client_close".into()))),
            ),
            DotnetClassExport::new(
                "dotnet.System.Net.Sockets",
                ClassType::new("TcpListener")
                    .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.tcp_listener_new"))
                    .with_method(MethodDef::new("Start",            0, MethodBody::Common("dotnet.tcp_listener_start".into())))
                    .with_method(MethodDef::new("Stop",             0, MethodBody::Common("dotnet.tcp_listener_stop".into())))
                    .with_method(MethodDef::new("AcceptTcpClient",  0, MethodBody::Common("dotnet.tcp_listener_accept".into())))
                    .with_method(MethodDef::new("Pending",          0, MethodBody::Common("dotnet.tcp_listener_pending".into()))),
            ),
            DotnetClassExport::new(
                "dotnet.System.Net.Sockets",
                ClassType::new("UdpClient")
                    .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.udp_client_new"))
                    .with_method(MethodDef::new("Send",    4, MethodBody::Common("dotnet.udp_send".into())))
                    .with_method(MethodDef::new("Receive", 0, MethodBody::Common("dotnet.udp_receive".into())))
                    .with_method(MethodDef::new("Close",   0, MethodBody::Common("dotnet.udp_close".into()))),
            ),
            DotnetClassExport::new(
                "dotnet.System.Net",
                ClassType::new("Dns")
                    .with_method(MethodDef::static_method("GetHostAddresses", 1, MethodBody::Common("dotnet.dns_get_host_addresses".into())))
                    .with_method(MethodDef::static_method("GetHostEntry",     1, MethodBody::Common("dotnet.dns_get_host_entry".into())))
                    .with_method(MethodDef::static_method("GetHostName",      0, MethodBody::Common("dotnet.dns_get_host_name".into()))),
            ),
            // .NET `System.IO.StreamReader` / `StreamWriter` — text I/O
            // wrappers, fully migrated to `emitter::dotnet::core::stream_io_adapter`.
            // Load-whole-file model via `node:fs.{read,write}FileSync`.
            // No `dotnet:io` host fns.
            DotnetClassExport::new(
                "dotnet.System.IO",
                ClassType::new("StreamReader")
                    .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.stream_reader_new"))
                    .with_method(MethodDef::new("ReadLine",  0, MethodBody::Common("dotnet.stream_reader_read_line".into())))
                    .with_method(MethodDef::new("ReadToEnd", 0, MethodBody::Common("dotnet.stream_reader_read_to_end".into())))
                    .with_method(MethodDef::new("Close",     0, MethodBody::Common("dotnet.stream_reader_close".into())))
                    .with_method(MethodDef::new("Dispose",   0, MethodBody::Common("dotnet.stream_reader_close".into()))),
            ),
            DotnetClassExport::new(
                "dotnet.System.IO",
                ClassType::new("StreamWriter")
                    .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.stream_writer_new"))
                    .with_method(MethodDef::new("Write",     1, MethodBody::Common("dotnet.stream_writer_write".into())))
                    .with_method(MethodDef::new("WriteLine", 1, MethodBody::Common("dotnet.stream_writer_write_line".into())))
                    .with_method(MethodDef::new("Flush",     0, MethodBody::Common("dotnet.stream_writer_flush".into())))
                    .with_method(MethodDef::new("Close",     0, MethodBody::Common("dotnet.stream_writer_flush".into())))
                    .with_method(MethodDef::new("Dispose",   0, MethodBody::Common("dotnet.stream_writer_flush".into()))),
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