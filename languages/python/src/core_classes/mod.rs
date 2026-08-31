//! Python stdlib classes the runtime provides, synthesized as ordinary AST.
//!
//! A builtin class is a CLASS. It is declared here as a `StmtKind::ClassDecl`
//! and appended to the module body, so it flows through the same path a user
//! class does — `normalize_class` → `NormalClass` → `compile_class` — and
//! inherits every piece of machinery that path already provides: a reserved
//! type slot and a real rtt at `struct.new_default $T`, the runtime
//! `TypeRegistry` registration, the prototype stamp that makes member dispatch
//! RECEIVER-based, MRO, and protocol-slot binding.
//!
//! ⛔ THIS REPLACES THE PRELUDES. `parse_python_prelude(X_PRELUDE)` splices
//! parsed Python SOURCE into the program body — 31 constants, 3,212 lines, 123
//! classes as of 2026-08-30. Declaring a class AS AST rather than as source
//! text is what keeps it free: no scan of the program text, no second parse,
//! and none of the prelude's failure modes (a silent `Vec::new()` on a parse
//! error, an indentation preprocessor that mis-reads `#` and `[` inside string
//! literals, a module member that resolves to nothing because the surface was
//! bare top-level `def`s).
//!
//! ⛔ And it replaces HAND-ROLLED construction. An adapter that builds an
//! object with `class_slots::emit_class_alloc` + stamped fields emits an
//! anonymous struct with no type, no vtable and no prototype — so `str(x)`
//! falls through to the object formatter and `type(x).__name__` answers
//! `object`. Both were measured on this very module before the rewrite.
//!
//! Bodies are deliberately plain: arithmetic, `str()`, `+`, `while`. Each
//! lowers through the shared machinery, so a class carries no python-private
//! emitter and the semantics are the ones every other language gets. Genuine
//! bit-twiddling stays in an emitter adapter behind a profile row, and a class
//! calling one is an ordinary call.
//!
//! Mirrors `languages/dart/src/core_classes/`.

mod builders;
mod contextlib;
mod csv;
mod http_ssl;
mod ipaddress;
mod logging;
mod futures;
mod multiprocessing;
mod pathlib;
mod queue;
mod subprocess;
mod threading;
mod typeobj;
mod time;
mod traceback;
mod warnings;

use vybe_ast::Statement;

/// Every class declared here, paired with its builder. The walker skips any
/// name the program declares itself, so a user `class IPv4Address` still wins.
///
/// Order matters where one class extends or constructs another: a class must
/// be declared after the one it depends on, so the ancestor's MRO is resolved
/// when the child's `__types` chain is stamped.
pub const CORE_CLASSES: &[(&str, fn() -> Statement)] = &[
    ("IPv4Address", ipaddress::ipv4_address),
    ("IPv4Network", ipaddress::ipv4_network),
    ("IPv4Interface", ipaddress::ipv4_interface),
    ("__WarningRecord", warnings::warning_record),
    ("__CatchWarnings", warnings::catch_warnings),
    ("LogRecord", logging::log_record),
    ("Formatter", logging::formatter),
    ("Filter", logging::filter_class),
    ("Handler", logging::handler),
    ("StreamHandler", logging::stream_handler),
    ("FileHandler", logging::file_handler),
    ("Logger", logging::logger),
    ("__NullContext", contextlib::null_context),
    ("__Closing", contextlib::closing),
    ("__Suppress", contextlib::suppress),
    ("__GenCM", contextlib::gen_cm),
    ("__Redirect", contextlib::redirect),
    ("FrameSummary", traceback::frame_summary),
    ("StackSummary", traceback::stack_summary),
    ("TracebackException", traceback::traceback_exception),
    ("HTTPMessage", http_ssl::http_message),
    ("HTTPResponse", http_ssl::http_response),
    ("HTTPConnection", http_ssl::http_connection),
    ("HTTPSConnection", http_ssl::https_connection),
    ("SSLContext", http_ssl::ssl_context),
    ("TLSVersion", http_ssl::tls_version),
    ("Purpose", http_ssl::purpose),
    ("__PyCsvExcel", csv::excel_dialect),
    ("__PyCsvReader", csv::reader),
    ("__PyCsvWriter", csv::writer),
    ("__PyCsvDictReader", csv::dict_reader),
    ("__PyCsvDictWriter", csv::dict_writer),
    ("Sniffer", csv::sniffer),
    ("__PyLock", threading::base_lock),
    ("Semaphore", threading::semaphore),
    ("BoundedSemaphore", threading::bounded_semaphore),
    ("Event", threading::event),
    ("Condition", threading::condition),
    ("Barrier", threading::barrier),
    ("local", threading::thread_local),
    ("Thread", threading::thread),
    ("Timer", threading::timer),
    ("Queue", queue::queue),
    ("LifoQueue", queue::lifo_queue),
    ("PriorityQueue", queue::priority_queue),
    ("SimpleQueue", queue::simple_queue),
    ("Future", futures::future),
    ("Process", multiprocessing::process),
    ("Pool", multiprocessing::pool),
    ("__PyValue", multiprocessing::shared_value),
    ("__PyProcessInfo", multiprocessing::process_info),
    ("Manager", multiprocessing::manager),
    ("__PyPipeEnd", multiprocessing::pipe_end),
    ("CompletedProcess", subprocess::completed_process),
    ("CalledProcessError", subprocess::called_process_error),
    ("TimeoutExpired", subprocess::timeout_expired),
    ("Popen", subprocess::popen),
    ("PurePath", pathlib::pure_path),
    ("Path", pathlib::path),
    ("__py_type_obj", typeobj::type_obj),
];

/// Classes generated from a table rather than a builder each — the warning
/// categories are eleven rows of `class X(Y): pass`, and writing eleven
/// builders would be eleven copies of one shape.
fn generated_classes(module: &str) -> Vec<Statement> {
    match module {
        "warnings" => warnings::CATEGORIES
            .iter()
            .map(|(name, parent)| warnings::category(name, parent))
            .collect(),
        // The http/ssl exception tree — five rows of `class X(Y): pass`, and
        // both modules pull the whole tree because `ssl.SSLError` extends
        // `OSError` while `http.client.HTTPException` extends `Exception`.
        // `Lock`/`RLock` are `__PyLock` with nothing added, so they are two
        // rows rather than two builders.
        // The two executors share one body; two rows, not two builders.
        "concurrent" => futures::EXECUTORS
            .iter()
            .map(|name| futures::executor(name))
            .collect(),
        // The two pure flavours differ only in `_is_win`.
        "pathlib" => pathlib::FLAVOURS
            .iter()
            .map(|(name, win)| pathlib::flavour(name, *win))
            .collect(),
        "queue" => queue::EXCEPTIONS
            .iter()
            .map(|(name, parent)| queue::exception(name, parent))
            .collect(),
        "threading" => threading::LOCK_ALIASES
            .iter()
            .map(|(name, parent)| threading::lock_alias(name, parent))
            .collect(),
        "http" | "ssl" => http_ssl::EXCEPTIONS
            .iter()
            .map(|(name, parent)| http_ssl::exception(name, parent))
            .collect(),
        _ => Vec::new(),
    }
}

/// Module-level functions declared alongside the classes, keyed by the MODULE
/// whose import pulls them in.
const MODULE_FUNCTIONS: &[(&str, fn() -> Vec<Statement>)] = &[
    ("ipaddress", ipaddress::module_functions),
    ("warnings", warnings::module_functions),
    ("logging", logging::module_functions),
    ("contextlib", contextlib::module_functions),
    ("traceback", traceback::module_functions),
    ("http", http_ssl::module_functions),
    ("csv", csv::module_functions),
    ("threading", threading::module_functions),
    
    ("concurrent", futures::module_functions),
    ("multiprocessing", multiprocessing::module_functions),
    ("subprocess", subprocess::module_functions),
    ("time", time::module_functions),
    ("pathlib", pathlib::module_functions),
];

/// The MODULE SURFACE: which `<module>.<name>` reads resolve to a declaration
/// made here, and under what global name.
///
/// This is the "registering the leaves" half. A declared class or function is
/// an ordinary global in the emitted module, so `ipaddress.ip_address(x)` is
/// `ip_address(x)` — the walker asks this table and rewrites the member read.
/// It is DATA, one row per exported name, not a per-module rewrite function.
pub const MODULE_SURFACE: &[(&str, &str, &str)] = &[
    ("ipaddress", "ip_address", "ip_address"),
    ("ipaddress", "ip_network", "IPv4Network"),
    ("ipaddress", "ip_interface", "IPv4Interface"),
    ("ipaddress", "collapse_addresses", "collapse_addresses"),
    ("ipaddress", "IPv4Address", "IPv4Address"),
    ("ipaddress", "IPv4Network", "IPv4Network"),
    ("ipaddress", "IPv4Interface", "IPv4Interface"),
    ("warnings", "warn", "warn"),
    ("warnings", "catch_warnings", "catch_warnings"),
    ("warnings", "filterwarnings", "filterwarnings"),
    ("warnings", "simplefilter", "simplefilter"),
    ("warnings", "resetwarnings", "resetwarnings"),
    ("warnings", "formatwarning", "formatwarning"),
    ("logging", "getLogger", "getLogger"),
    ("logging", "basicConfig", "basicConfig"),
    ("logging", "getLevelName", "getLevelName"),
    ("logging", "addLevelName", "addLevelName"),
    ("logging", "debug", "debug"),
    ("logging", "info", "info"),
    ("logging", "warning", "warning"),
    ("logging", "error", "error"),
    ("logging", "critical", "critical"),
    ("logging", "log", "log"),
    ("logging", "LogRecord", "LogRecord"),
    ("logging", "Formatter", "Formatter"),
    ("logging", "Filter", "Filter"),
    ("logging", "Handler", "Handler"),
    ("logging", "StreamHandler", "StreamHandler"),
    ("logging", "FileHandler", "FileHandler"),
    ("logging", "Logger", "Logger"),
    ("contextlib", "nullcontext", "nullcontext"),
    ("contextlib", "closing", "closing"),
    ("contextlib", "redirect_stdout", "redirect_stdout"),
    ("contextlib", "redirect_stderr", "redirect_stderr"),
    ("contextlib", "suppress", "suppress"),
    ("contextlib", "contextmanager", "contextmanager"),
    ("http.client", "HTTPConnection", "HTTPConnection"),
    ("http.client", "HTTPSConnection", "HTTPSConnection"),
    ("http.client", "HTTPResponse", "HTTPResponse"),
    ("http.client", "HTTPMessage", "HTTPMessage"),
    ("http.client", "HTTPException", "HTTPException"),
    ("http.client", "BadStatusLine", "BadStatusLine"),
    ("http.client", "IncompleteRead", "IncompleteRead"),
    ("http.client", "parse_headers", "parse_headers"),
    ("ssl", "SSLContext", "SSLContext"),
    ("ssl", "SSLError", "SSLError"),
    ("ssl", "CertificateError", "CertificateError"),
    ("ssl", "TLSVersion", "TLSVersion"),
    ("ssl", "Purpose", "Purpose"),
    ("ssl", "create_default_context", "create_default_context"),
    ("ssl", "match_hostname", "match_hostname"),
    ("ssl", "enum_certificates", "enum_certificates"),
    ("ssl", "wrap_socket", "ssl_wrap_socket"),
    ("csv", "reader", "reader"),
    ("csv", "writer", "writer"),
    ("csv", "DictReader", "DictReader"),
    ("csv", "DictWriter", "DictWriter"),
    ("csv", "Sniffer", "Sniffer"),
    ("csv", "excel", "excel"),
    ("csv", "list_dialects", "list_dialects"),
    ("csv", "field_size_limit", "field_size_limit"),
    ("threading", "Lock", "Lock"),
    ("threading", "RLock", "RLock"),
    ("threading", "Semaphore", "Semaphore"),
    ("threading", "BoundedSemaphore", "BoundedSemaphore"),
    ("threading", "Event", "Event"),
    ("threading", "Condition", "Condition"),
    ("threading", "Barrier", "Barrier"),
    ("threading", "local", "local"),
    ("threading", "Thread", "Thread"),
    ("threading", "Timer", "Timer"),
    ("queue", "Queue", "Queue"),
    ("queue", "LifoQueue", "LifoQueue"),
    ("queue", "PriorityQueue", "PriorityQueue"),
    ("queue", "SimpleQueue", "SimpleQueue"),
    ("queue", "Empty", "Empty"),
    ("queue", "Full", "Full"),
    ("concurrent.futures", "Future", "Future"),
    ("concurrent.futures", "ThreadPoolExecutor", "ThreadPoolExecutor"),
    ("concurrent.futures", "ProcessPoolExecutor", "ProcessPoolExecutor"),
    ("concurrent.futures", "as_completed", "as_completed"),
    ("concurrent.futures", "wait", "wait"),
    ("multiprocessing", "Process", "Process"),
    ("multiprocessing", "Pool", "Pool"),
    ("multiprocessing", "Manager", "Manager"),
    ("multiprocessing", "Value", "Value"),
    ("multiprocessing", "Array", "Array"),
    ("multiprocessing", "Pipe", "Pipe"),
    ("multiprocessing", "cpu_count", "cpu_count"),
    ("multiprocessing", "current_process", "current_process"),
    ("multiprocessing", "active_children", "active_children"),
    ("subprocess", "CompletedProcess", "CompletedProcess"),
    ("subprocess", "CalledProcessError", "CalledProcessError"),
    ("subprocess", "TimeoutExpired", "TimeoutExpired"),
    ("subprocess", "Popen", "Popen"),
    ("subprocess", "run", "run"),
    ("subprocess", "call", "call"),
    ("subprocess", "check_output", "check_output"),
    ("subprocess", "check_call", "check_call"),
    ("threading", "current_thread", "current_thread"),
    ("threading", "main_thread", "main_thread"),
    ("threading", "enumerate", "enumerate"),
    ("threading", "active_count", "active_count"),
    ("threading", "get_ident", "get_ident"),
    ("threading", "stack_size", "stack_size"),
    ("subprocess", "PIPE", "PIPE"),
    ("subprocess", "STDOUT", "STDOUT"),
    ("subprocess", "DEVNULL", "DEVNULL"),
    ("threading", "excepthook", "excepthook"),
    ("time", "sleep", "sleep"),
    ("pathlib", "PurePath", "PurePath"),
    ("pathlib", "PurePosixPath", "PurePosixPath"),
    ("pathlib", "PureWindowsPath", "PureWindowsPath"),
    ("pathlib", "Path", "Path"),







    ("traceback", "format_exc", "format_exc"),
    ("traceback", "format_exception", "format_exception"),
    ("traceback", "format_exception_only", "format_exception_only"),
    ("traceback", "format_tb", "format_tb"),
    ("traceback", "format_stack", "format_stack"),
    ("traceback", "extract_tb", "extract_tb"),
    ("traceback", "extract_stack", "extract_stack"),
    ("traceback", "print_exc", "print_exc"),
    ("traceback", "print_tb", "print_tb"),
    ("traceback", "print_stack", "print_stack"),
    ("traceback", "print_exception", "print_exception"),
    ("traceback", "clear_frames", "clear_frames"),
    ("traceback", "walk_tb", "walk_tb"),
    ("traceback", "walk_stack", "walk_stack"),
    ("traceback", "FrameSummary", "FrameSummary"),
    ("traceback", "StackSummary", "StackSummary"),
    ("traceback", "TracebackException", "TracebackException"),
];

/// The PROPERTY names each declared class exposes, for the walker's
/// `py_class_properties` registry.
///
/// ⛔ A spliced class is never walked, so nothing calls
/// `note_class_property_kind` for it — and a property the walker does not know
/// about is read as a plain attribute: `PurePath("a/b").drive` resolved to an
/// absent field instead of invoking the accessor. Fields worked and getters did
/// not, which is exactly the shape that says "the registry, not the AST".
pub const CLASS_PROPERTIES: &[(&str, &[&str])] = &[
    (
        "PurePath",
        &["drive", "root", "anchor", "name", "stem", "suffix", "suffixes",
          "parts", "parent"],
    ),
    ("PurePosixPath", &["drive", "root", "anchor", "name", "stem", "suffix",
                        "suffixes", "parts", "parent"]),
    ("PureWindowsPath", &["drive", "root", "anchor", "name", "stem", "suffix",
                          "suffixes", "parts", "parent"]),
    (
        "Path",
        &["drive", "root", "anchor", "name", "stem", "suffix", "suffixes",
          "parts", "parent"],
    ),
];

/// The class names a program actually needs, on the same gate as
/// `declarations_for`.
///
/// ⛔ The walker must register ONLY these. Registering every core class name in
/// `py_defined_classes` for every program cost **242 tests** — sets,
/// comprehensions and zip — because that registry is consulted far more widely
/// than "is this name a class", and seeding it changes lowering decisions in
/// programs that never mention the module.
pub fn needed_classes(source: &str) -> Vec<&'static str> {
    let mut out = Vec::new();
    for (module, classes) in MODULE_CLASSES {
        if source.contains(module) {
            out.extend_from_slice(classes);
            if *module == "warnings" {
                out.extend(warnings::CATEGORIES.iter().map(|(name, _)| *name));
            }
            if *module == "threading" {
                out.extend(threading::LOCK_ALIASES.iter().map(|(name, _)| *name));
            }
            if *module == "queue" {
                out.extend(queue::EXCEPTIONS.iter().map(|(name, _)| *name));
            }
            if *module == "pathlib" {
                out.extend(pathlib::FLAVOURS.iter().map(|(name, _)| *name));
            }
            if *module == "concurrent" {
                out.extend(futures::EXECUTORS.iter().copied());
            }
            if *module == "http" || *module == "ssl" {
                out.extend(http_ssl::EXCEPTIONS.iter().map(|(name, _)| *name));
            }
        }
    }
    out
}

/// The instance attributes each declared class carries, for the walker's
/// `py_class_attrs` registry. The walker decides how to lower `a.attr` from
/// this — a class it has no attrs for takes the untyped path — so a declared
/// class has to state them exactly as the walk would have collected them.
pub const CLASS_ATTRS: &[(&str, &[&str])] = &[
    (
        "IPv4Address",
        &[
            "version", "_int", "_text", "compressed", "exploded", "packed",
            "is_private", "is_loopback", "is_multicast", "is_global",
        ],
    ),
    (
        "IPv4Network",
        &[
            "version", "prefixlen", "num_addresses", "_base", "network_address",
            "netmask", "hostmask", "broadcast_address",
        ],
    ),
    ("IPv4Interface", &["version", "prefixlen", "ip", "network"]),
];

/// The global a `<module>.<name>` read denotes, if this module declares it.
pub fn module_member(module: &str, name: &str) -> Option<&'static str> {
    MODULE_SURFACE
        .iter()
        .find(|(m, n, _)| *m == module && *n == name)
        .map(|(_, _, global)| *global)
}

/// Which classes a module's import needs. A class is not reachable by its own
/// name in Python — `ipaddress.IPv4Address` is, but a program far more often
/// only ever names `ip_address` — so the gate is the MODULE, not the class.
const MODULE_CLASSES: &[(&str, &[&str])] = &[
    ("ipaddress", &["IPv4Address", "IPv4Network", "IPv4Interface"]),
    ("warnings", &["__WarningRecord", "__CatchWarnings"]),
    (
        "logging",
        &["LogRecord", "Formatter", "Filter", "Handler", "StreamHandler",
          "FileHandler", "Logger"],
    ),
    (
        "contextlib",
        &["__NullContext", "__Closing", "__Suppress", "__GenCM", "__Redirect"],
    ),
    ("traceback", &["FrameSummary", "StackSummary", "TracebackException"]),
    (
        "http",
        &["HTTPMessage", "HTTPResponse", "HTTPConnection", "HTTPSConnection"],
    ),
    ("ssl", &["SSLContext", "TLSVersion", "Purpose"]),
    (
        "csv",
        &["__PyCsvExcel", "__PyCsvReader", "__PyCsvWriter", "__PyCsvDictReader",
          "__PyCsvDictWriter", "Sniffer"],
    ),
    (
        "threading",
        &["__PyLock", "Semaphore", "BoundedSemaphore", "Event", "Condition",
          "Barrier", "local", "Thread", "Timer"],
    ),
    (
        "queue",
        &["Queue", "LifoQueue", "PriorityQueue", "SimpleQueue"],
    ),
    ("concurrent", &["Future"]),
    (
        "multiprocessing",
        &["Process", "Pool", "__PyValue", "__PyProcessInfo", "Manager",
          "__PyPipeEnd"],
    ),
    (
        "subprocess",
        &["CompletedProcess", "CalledProcessError", "TimeoutExpired", "Popen"],
    ),
    // No classes — the entry exists so `time`'s module FUNCTIONS splice.
    ("time", &[]),
    ("pathlib", &["PurePath", "Path"]),
    // Not a module — the gate is the ANNOTATION machinery that builds it.
    ("__annotations__", &["__py_type_obj"]),
];

/// The declarations a program needs, given its source text.
///
/// **The substring gate is deliberately CONSERVATIVE**, exactly as dart's is: a
/// module name inside a comment declares classes nobody uses, which costs
/// compile time and nothing else. It cannot go the other way — a program that
/// imports the module has the name in its text by definition.
///
/// It is NOT the prelude's gate, despite the same shape. A prelude's cost is a
/// full parse of hundreds of lines of Python; this is a `Vec<Statement>` built
/// by function calls.
pub fn declarations_for(source: &str, is_user_declared: impl Fn(&str) -> bool) -> Vec<Statement> {
    let mut out = Vec::new();
    for (module, classes) in MODULE_CLASSES {
        if !source.contains(module) {
            continue;
        }
        out.extend(generated_classes(module));
        for (name, build) in CORE_CLASSES {
            if classes.contains(name) && !is_user_declared(name) {
                out.push(build());
            }
        }
        for (owner, build) in MODULE_FUNCTIONS {
            if owner == module {
                out.extend(build());
            }
        }
    }
    out
}
