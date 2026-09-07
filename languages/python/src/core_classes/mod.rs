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
mod argparse;
mod ast_mod;
mod builtins;
mod bytes;
mod collections;
mod configparser;
mod contextlib;
mod contextvars;
mod csv;
mod decimal;
mod difflib;
mod exceptions;
mod filecmp;
mod fnmatch;
mod fractions;
mod futures;
mod gettext;
mod graphlib;
mod html_parser;
mod http_ssl;
mod io;
mod ipaddress;
mod json_mod;
mod linecache;
mod logging;
mod mock;
mod multiprocessing;
mod object_class;
mod optparse;
mod pathlib;
mod plistlib;
mod pprint;
mod queue;
mod random_mod;
mod sched;
mod selectors;
mod shlex;
mod shutil;
mod site;
mod socket;
mod string_mod;
mod struct_mod;
mod subprocess;
mod sysconfig;
mod tempfile;
mod textwrap;
mod threading;
mod time;
mod timeit;
mod tokenize;
mod tomllib;
mod traceback;
mod tracemalloc;
mod types_mod;
mod typeobj;
mod warnings;

use vybe_ast::Statement;

/// Every class declared here, paired with its builder. The walker skips any
/// name the program declares itself, so a user `class IPv4Address` still wins.
///
/// Order matters where one class extends or constructs another: a class must
/// be declared after the one it depends on, so the ancestor's MRO is resolved
/// when the child's `__types` chain is stamped.
pub const CORE_CLASSES: &[(&str, fn() -> Statement)] = &[
    ("ellipsis", builtins::ellipsis_type),
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
    ("__AsyncGenCM", contextlib::async_gen_cm),
    ("__Redirect", contextlib::redirect),
    ("Namespace", argparse::namespace),
    ("ArgumentParser", argparse::argument_parser),
    ("__ArgparseGroup", argparse::argparse_group),
    ("__ArgparseSubparsers", argparse::argparse_subparsers),
    ("Module", ast_mod::module_node),
    ("Expression", ast_mod::expression_node),
    ("Constant", ast_mod::constant_node),
    ("Assign", || ast_mod::simple_node("Assign")),
    ("Name", || ast_mod::simple_node("Name")),
    ("BinOp", || ast_mod::simple_node("BinOp")),
    ("Add", || ast_mod::simple_node("Add")),
    ("FunctionDef", || ast_mod::simple_node("FunctionDef")),
    ("Return", || ast_mod::simple_node("Return")),
    ("NodeVisitor", ast_mod::node_visitor),
    ("NodeTransformer", ast_mod::node_transformer),
    ("JSONDecoder", json_mod::json_decoder),
    ("JSONEncoder", json_mod::json_encoder),
    ("SelectorKey", selectors::selector_key),
    ("SelectSelector", selectors::select_selector),
    ("EpollSelector", selectors::epoll_selector),
    ("KqueueSelector", selectors::kqueue_selector),
    ("PollSelector", selectors::poll_selector),
    ("DevpollSelector", selectors::devpoll_selector),
    ("FrameSummary", traceback::frame_summary),
    ("StackSummary", traceback::stack_summary),
    ("TracebackException", traceback::traceback_exception),
    ("__PyTraceFrame", tracemalloc::frame),
    ("__PyTraceback", tracemalloc::trace),
    ("__PyTraceStat", tracemalloc::statistic),
    ("__PyTraceStatDiff", tracemalloc::statistic_diff),
    ("Snapshot", tracemalloc::snapshot),
    ("Filter", tracemalloc::filter),
    ("HTTPMessage", http_ssl::http_message),
    ("HTTPResponse", http_ssl::http_response),
    ("HTTPConnection", http_ssl::http_connection),
    ("HTTPSConnection", http_ssl::https_connection),
    ("HTTPStatus", http_ssl::http_status),
    ("Morsel", http_ssl::morsel),
    ("SimpleCookie", http_ssl::simple_cookie),
    ("CookieJar", http_ssl::cookie_jar),
    ("LWPCookieJar", http_ssl::lwp_cookie_jar),
    ("Request", http_ssl::urllib_request),
    ("HTMLParser", html_parser::html_parser),
    ("SSLContext", http_ssl::ssl_context),
    ("TLSVersion", http_ssl::tls_version),
    ("Purpose", http_ssl::purpose),
    ("__PyCsvExcel", csv::excel_dialect),
    ("__PyCsvExcelTab", csv::excel_tab_dialect),
    ("__PyCsvSemi", csv::semi_dialect),
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
    ("__TimeitTimer", timeit::timer_class),
    ("TokenInfo", tokenize::token_info),
    ("TokenError", tokenize::token_error),
    ("__PyDirCmp", filecmp::dircmp),
    ("NullTranslations", gettext::null_translations),
    ("GNUTranslations", gettext::gnu_translations),
    ("__PyOptValues", optparse::values),
    ("OptionParser", optparse::option_parser),
    ("OptionGroup", optparse::option_group),
    ("SchedEvent", sched::event),
    ("scheduler", sched::scheduler),
    ("Token", contextvars::token),
    ("ContextVar", contextvars::context_var),
    ("PurePath", pathlib::pure_path),
    ("Path", pathlib::path),
    ("__PyNamedTempFile", tempfile::named_temp_file),
    ("TopologicalSorter", graphlib::topological_sorter),
    ("__pprint_PrettyPrinter", pprint::pretty_printer),
    ("__py_shlex_class", shlex::shlex_class),
    ("__py_TextWrapper", textwrap::text_wrapper),
    ("__string_Template", string_mod::template),
    ("__string_Formatter", string_mod::formatter),
    ("VybeSocketImpl", socket::socket_impl),
    ("ConfigParser", configparser::config_parser),
    ("RawConfigParser", configparser::raw_config_parser),
    ("BasicInterpolation", configparser::basic_interpolation),
    (
        "ExtendedInterpolation",
        configparser::extended_interpolation,
    ),
    ("StringIO", io::string_io),
    ("BytesIO", io::bytes_io),
    ("__PyDiskUsage", shutil::disk_usage_result),
    ("__PyTerminalSize", shutil::terminal_size),
    ("Context", decimal::context),
    ("DecimalTuple", decimal::decimal_tuple),
    ("Decimal", decimal::decimal),
    ("__PyDiffMatch", difflib::match_class),
    ("SequenceMatcher", difflib::sequence_matcher),
    ("Differ", difflib::differ),
    ("HtmlDiff", difflib::html_diff),
    ("__PyMockAny", mock::any),
    ("__PyMockCall", mock::call_record),
    ("__PyMockCallFactory", mock::call_factory),
    ("__PyMockNamedCallFactory", mock::named_call_factory),
    ("Mock", mock::mock),
    ("MagicMock", mock::magic_mock),
    ("PropertyMock", mock::property_mock),
    ("__PyPatch", mock::patch_context),
    ("__PyPatchDict", mock::patch_dict_context),
    ("__PyPatchFactory", mock::patch_factory),
    ("TOMLDecodeError", tomllib::toml_decode_error),
    ("Fraction", fractions::fraction),
    ("__py_type_obj", typeobj::type_obj),
    ("__py_SystemRandom", random_mod::system_random),
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
        "decimal" => decimal::EXCEPTIONS
            .iter()
            .map(|(name, parent)| decimal::exception(name, parent))
            .collect(),
        "graphlib" => graphlib::EXCEPTIONS
            .iter()
            .map(|(name, parent)| graphlib::exception(name, parent))
            .collect(),
        "socket" => socket::EXCEPTIONS
            .iter()
            .map(|(name, parent)| socket::exception(name, parent))
            .collect(),
        "queue" => queue::EXCEPTIONS
            .iter()
            .map(|(name, parent)| queue::exception(name, parent))
            .collect(),
        "threading" => threading::LOCK_ALIASES
            .iter()
            .map(|(name, parent)| threading::lock_alias(name, parent))
            .collect(),
        "contextvars" => vec![contextvars::context()],
        "configparser" => configparser::EXCEPTIONS
            .iter()
            .map(|(name, parent)| configparser::exception(name, parent))
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
    ("argparse", argparse::module_functions),
    ("ast", ast_mod::module_functions),
    ("plistlib", plistlib::module_functions),
    ("selectors", selectors::module_functions),
    ("traceback", traceback::module_functions),
    ("http", http_ssl::module_functions),
    ("ssl", http_ssl::module_functions),
    ("csv", csv::module_functions),
    ("threading", threading::module_functions),
    ("concurrent", futures::module_functions),
    ("multiprocessing", multiprocessing::module_functions),
    ("subprocess", subprocess::module_functions),
    ("time", time::module_functions),
    ("timeit", timeit::module_functions),
    ("tokenize", tokenize::module_functions),
    ("tomllib", tomllib::module_functions),
    ("tracemalloc", tracemalloc::module_functions),
    ("pathlib", pathlib::module_functions),
    ("fractions", fractions::module_functions),
    ("decimal", decimal::module_functions),
    ("filecmp", filecmp::module_functions),
    ("difflib", difflib::module_functions),
    ("unittest.mock", mock::module_functions),
    ("linecache", linecache::module_functions),
    ("gettext", gettext::module_functions),
    ("optparse", optparse::module_functions),
    ("sched", sched::module_functions),
    ("contextvars", contextvars::module_functions),
    ("shutil", shutil::module_functions),
    ("site", site::module_functions),
    ("io", io::module_functions),
    ("collections", collections_functions),
    ("fnmatch", fnmatch::module_functions),
    ("socket", socket::module_functions),
    ("pprint", pprint::module_functions),
    ("shlex", shlex::module_functions),
    ("tempfile", tempfile::module_functions),
    ("struct", struct_mod::module_functions),
    ("sysconfig", sysconfig::module_functions),
    ("types", types_mod::module_functions),
];

/// The MODULE SURFACE: which `<module>.<name>` reads resolve to a declaration
/// made here, and under what global name.
///
/// This is the "registering the leaves" half. A declared class or function is
/// an ordinary global in the emitted module, so `ipaddress.ip_address(x)` is
/// `ip_address(x)` — the walker asks this table and rewrites the member read.
/// It is DATA, one row per exported name, not a per-module rewrite function.
pub const MODULE_SURFACE: &[(&str, &str, &str)] = &[
    ("random", "SystemRandom", "__py_SystemRandom"),
    ("fractions", "Fraction", "Fraction"),
    ("shutil", "copyfile", "copyfile"),
    ("shutil", "copy", "copy"),
    ("shutil", "copy2", "copy2"),
    ("shutil", "copymode", "copymode"),
    ("shutil", "copystat", "copystat"),
    ("shutil", "chown", "chown"),
    ("shutil", "copyfileobj", "copyfileobj"),
    ("shutil", "move", "move"),
    ("shutil", "rmtree", "rmtree"),
    ("shutil", "copytree", "copytree"),
    ("shutil", "which", "which"),
    ("shutil", "disk_usage", "disk_usage"),
    ("shutil", "get_terminal_size", "get_terminal_size"),
    ("shutil", "get_archive_formats", "get_archive_formats"),
    ("shutil", "get_unpack_formats", "get_unpack_formats"),
    ("shutil", "make_archive", "make_archive"),
    ("shutil", "unpack_archive", "unpack_archive"),
    ("shutil", "ignore_patterns", "ignore_patterns"),
    ("struct", "pack", "__py_struct_pack"),
    ("struct", "unpack", "__py_struct_unpack"),
    ("struct", "calcsize", "__py_struct_calcsize"),
    ("timeit", "timeit", "timeit"),
    ("timeit", "repeat", "repeat"),
    ("timeit", "default_timer", "default_timer"),
    ("timeit", "Timer", "__TimeitTimer"),
    ("filecmp", "cmp", "cmp"),
    ("filecmp", "cmpfiles", "cmpfiles"),
    ("filecmp", "dircmp", "dircmp"),
    ("filecmp", "clear_cache", "clear_cache"),
    ("difflib", "SequenceMatcher", "SequenceMatcher"),
    ("difflib", "Differ", "Differ"),
    ("difflib", "HtmlDiff", "HtmlDiff"),
    ("difflib", "get_close_matches", "get_close_matches"),
    ("difflib", "restore", "restore"),
    ("difflib", "unified_diff", "unified_diff"),
    ("difflib", "context_diff", "context_diff"),
    ("difflib", "IS_CHARACTER_JUNK", "IS_CHARACTER_JUNK"),
    ("difflib", "IS_LINE_JUNK", "IS_LINE_JUNK"),
    ("unittest.mock", "Mock", "Mock"),
    ("unittest.mock", "MagicMock", "MagicMock"),
    ("unittest.mock", "PropertyMock", "PropertyMock"),
    ("unittest.mock", "ANY", "ANY"),
    ("unittest.mock", "call", "call"),
    ("unittest.mock", "patch", "patch"),
    ("unittest.mock", "seal", "seal"),
    ("unittest.mock", "sys", "sys"),
    ("linecache", "cache", "__py_linecache_cache"),
    ("linecache", "getline", "__py_linecache_getline"),
    ("linecache", "getlines", "__py_linecache_getlines"),
    ("linecache", "clearcache", "__py_linecache_clearcache"),
    ("linecache", "checkcache", "__py_linecache_checkcache"),
    ("linecache", "updatecache", "__py_linecache_updatecache"),
    ("linecache", "lazycache", "__py_linecache_lazycache"),
    ("gettext", "NullTranslations", "NullTranslations"),
    ("gettext", "GNUTranslations", "GNUTranslations"),
    ("gettext", "gettext", "__py_gettext_gettext"),
    ("gettext", "dgettext", "__py_gettext_dgettext"),
    ("gettext", "ngettext", "__py_gettext_ngettext"),
    ("gettext", "dngettext", "__py_gettext_dngettext"),
    ("gettext", "pgettext", "__py_gettext_pgettext"),
    ("gettext", "dpgettext", "__py_gettext_dpgettext"),
    ("gettext", "npgettext", "__py_gettext_npgettext"),
    ("gettext", "dnpgettext", "__py_gettext_dnpgettext"),
    ("gettext", "bindtextdomain", "__py_gettext_bindtextdomain"),
    ("gettext", "textdomain", "__py_gettext_textdomain"),
    ("gettext", "find", "__py_gettext_find"),
    ("gettext", "translation", "__py_gettext_translation"),
    ("gettext", "install", "__py_gettext_install"),
    ("optparse", "OptionParser", "OptionParser"),
    ("optparse", "OptionGroup", "OptionGroup"),
    ("sched", "Event", "SchedEvent"),
    ("sched", "scheduler", "scheduler"),
    ("contextvars", "ContextVar", "ContextVar"),
    ("contextvars", "Token", "Token"),
    ("contextvars", "Context", "Context"),
    ("contextvars", "copy_context", "copy_context"),
    ("tomllib", "TOMLDecodeError", "TOMLDecodeError"),
    ("logging", "exception", "exception"),
    ("logging", "setLogRecordFactory", "setLogRecordFactory"),
    ("logging", "getLogRecordFactory", "getLogRecordFactory"),
    ("logging.config", "dictConfig", "dictConfig"),
    ("logging.config", "fileConfig", "fileConfig"),
    (
        "logging.handlers",
        "RotatingFileHandler",
        "RotatingFileHandler",
    ),
    ("warnings", "showwarning", "showwarning"),
    ("warnings", "warn_explicit", "warn_explicit"),
    ("warnings", "_filters_mutated", "_filters_mutated"),
    ("warnings", "onceregistry", "__py_warnings_onceregistry"),
    ("tokenize", "generate_tokens", "generate_tokens"),
    ("tokenize", "tokenize", "tokenize"),
    ("tokenize", "untokenize", "untokenize"),
    ("tokenize", "detect_encoding", "detect_encoding"),
    ("tokenize", "TokenInfo", "TokenInfo"),
    ("tokenize", "TokenError", "TokenError"),
    ("tempfile", "NamedTemporaryFile", "__PyNamedTempFile"),
    ("tempfile", "TemporaryFile", "__PyNamedTempFile"),
    ("graphlib", "TopologicalSorter", "TopologicalSorter"),
    ("graphlib", "CycleError", "CycleError"),
    ("pprint", "pprint", "__pprint_pprint"),
    ("pprint", "pformat", "__pprint_pformat"),
    ("pprint", "pp", "__pprint_pp"),
    ("pprint", "saferepr", "__pprint_saferepr"),
    ("pprint", "isrecursive", "__pprint_isrecursive"),
    ("pprint", "isreadable", "__pprint_isreadable"),
    ("pprint", "PrettyPrinter", "__pprint_PrettyPrinter"),
    ("shlex", "split", "__py_shlex_split"),
    ("shlex", "quote", "__py_shlex_quote"),
    ("shlex", "join", "__py_shlex_join"),
    ("shlex", "shlex", "__py_shlex_class"),
    ("textwrap", "TextWrapper", "__py_TextWrapper"),
    ("string", "Template", "__string_Template"),
    ("string", "Formatter", "__string_Formatter"),
    ("socket", "socket", "VybeSocketImpl"),
    ("socket", "timeout", "VybeSocketTimeout"),
    ("socket", "gaierror", "VybeSocketGaiError"),
    ("socket", "create_connection", "create_connection"),
    ("socket", "inet_pton", "inet_pton"),
    ("socket", "inet_ntop", "inet_ntop"),
    ("socket", "ntohl", "ntohl"),
    ("socket", "htonl", "htonl"),
    ("socket", "ntohs", "ntohs"),
    ("socket", "htons", "htons"),
    ("socket", "getdefaulttimeout", "getdefaulttimeout"),
    ("socket", "setdefaulttimeout", "setdefaulttimeout"),
    ("fnmatch", "fnmatch", "__py_fnmatch_match"),
    ("fnmatch", "fnmatchcase", "__py_fnmatch_match"),
    ("fnmatch", "filter", "__py_fnmatch_filter"),
    ("fnmatch", "translate", "__fn_translate"),
    ("configparser", "ConfigParser", "ConfigParser"),
    ("configparser", "RawConfigParser", "RawConfigParser"),
    (
        "configparser",
        "BasicInterpolation",
        "BasicInterpolation",
    ),
    (
        "configparser",
        "ExtendedInterpolation",
        "ExtendedInterpolation",
    ),
    ("configparser", "Error", "Error"),
    (
        "configparser",
        "DuplicateSectionError",
        "DuplicateSectionError",
    ),
    ("configparser", "NoSectionError", "NoSectionError"),
    ("configparser", "NoOptionError", "NoOptionError"),
    ("io", "StringIO", "StringIO"),
    ("io", "BytesIO", "BytesIO"),
    ("decimal", "Decimal", "Decimal"),
    ("decimal", "Context", "Context"),
    ("decimal", "DecimalTuple", "DecimalTuple"),
    ("decimal", "getcontext", "getcontext"),
    ("decimal", "setcontext", "setcontext"),
    ("decimal", "localcontext", "localcontext"),
    ("decimal", "ROUND_HALF_UP", "ROUND_HALF_UP"),
    ("decimal", "ROUND_HALF_EVEN", "ROUND_HALF_EVEN"),
    ("decimal", "ROUND_HALF_DOWN", "ROUND_HALF_DOWN"),
    ("decimal", "ROUND_UP", "ROUND_UP"),
    ("decimal", "ROUND_DOWN", "ROUND_DOWN"),
    ("decimal", "ROUND_CEILING", "ROUND_CEILING"),
    ("decimal", "ROUND_FLOOR", "ROUND_FLOOR"),
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
    ("contextlib", "asynccontextmanager", "asynccontextmanager"),
    ("argparse", "ArgumentParser", "ArgumentParser"),
    ("argparse", "Namespace", "Namespace"),
    ("ast", "parse", "parse"),
    ("ast", "dump", "dump"),
    ("ast", "unparse", "unparse"),
    ("ast", "walk", "walk"),
    ("ast", "iter_child_nodes", "iter_child_nodes"),
    ("ast", "iter_fields", "iter_fields"),
    ("ast", "fix_missing_locations", "fix_missing_locations"),
    ("ast", "get_docstring", "get_docstring"),
    ("ast", "literal_eval", "literal_eval"),
    ("ast", "Module", "Module"),
    ("ast", "Expression", "Expression"),
    ("ast", "Constant", "Constant"),
    ("ast", "Assign", "Assign"),
    ("ast", "Name", "Name"),
    ("ast", "BinOp", "BinOp"),
    ("ast", "Add", "Add"),
    ("ast", "FunctionDef", "FunctionDef"),
    ("ast", "Return", "Return"),
    ("ast", "NodeVisitor", "NodeVisitor"),
    ("ast", "NodeTransformer", "NodeTransformer"),
    ("plistlib", "dumps", "dumps"),
    ("plistlib", "loads", "loads"),
    ("json", "JSONDecoder", "JSONDecoder"),
    ("json", "JSONEncoder", "JSONEncoder"),
    ("selectors", "DefaultSelector", "SelectSelector"),
    ("selectors", "SelectSelector", "SelectSelector"),
    ("selectors", "EpollSelector", "EpollSelector"),
    ("selectors", "KqueueSelector", "KqueueSelector"),
    ("selectors", "PollSelector", "PollSelector"),
    ("selectors", "DevpollSelector", "DevpollSelector"),
    ("selectors", "SelectorKey", "SelectorKey"),
    ("selectors", "EVENT_READ", "EVENT_READ"),
    ("selectors", "EVENT_WRITE", "EVENT_WRITE"),
    ("http.client", "HTTPConnection", "HTTPConnection"),
    ("http.client", "HTTPSConnection", "HTTPSConnection"),
    ("http.client", "HTTPResponse", "HTTPResponse"),
    ("http.client", "HTTPMessage", "HTTPMessage"),
    ("http.client", "HTTPException", "HTTPException"),
    ("http.client", "BadStatusLine", "BadStatusLine"),
    ("http.client", "IncompleteRead", "IncompleteRead"),
    ("http.client", "CannotSendRequest", "CannotSendRequest"),
    ("http.client", "parse_headers", "parse_headers"),
    ("http.client", "responses", "responses"),
    ("http", "HTTPStatus", "HTTPStatus"),
    ("http.cookies", "Morsel", "Morsel"),
    ("http.cookies", "SimpleCookie", "SimpleCookie"),
    ("http.cookies", "CookieError", "CookieError"),
    ("http.cookiejar", "CookieJar", "CookieJar"),
    ("http.cookiejar", "LWPCookieJar", "LWPCookieJar"),
    ("urllib.request", "Request", "Request"),
    ("html.parser", "HTMLParser", "HTMLParser"),
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
    ("csv", "excel_tab", "excel_tab"),
    ("csv", "list_dialects", "list_dialects"),
    ("csv", "field_size_limit", "field_size_limit"),
    ("csv", "get_dialect", "get_dialect"),
    ("csv", "register_dialect", "register_dialect"),
    ("csv", "unregister_dialect", "unregister_dialect"),
    ("csv", "QUOTE_MINIMAL", "QUOTE_MINIMAL"),
    ("csv", "QUOTE_ALL", "QUOTE_ALL"),
    ("csv", "QUOTE_NONNUMERIC", "QUOTE_NONNUMERIC"),
    ("csv", "QUOTE_NONE", "QUOTE_NONE"),
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
    (
        "concurrent.futures",
        "ThreadPoolExecutor",
        "ThreadPoolExecutor",
    ),
    (
        "concurrent.futures",
        "ProcessPoolExecutor",
        "ProcessPoolExecutor",
    ),
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
    ("sysconfig", "get_python_version", "get_python_version"),
    ("sysconfig", "get_platform", "get_platform"),
    ("sysconfig", "is_python_build", "is_python_build"),
    ("sysconfig", "get_default_scheme", "get_default_scheme"),
    ("sysconfig", "get_scheme_names", "get_scheme_names"),
    ("sysconfig", "get_path_names", "get_path_names"),
    ("sysconfig", "get_paths", "get_paths"),
    ("sysconfig", "get_path", "get_path"),
    ("sysconfig", "get_config_var", "get_config_var"),
    ("sysconfig", "get_config_vars", "get_config_vars"),
    ("sysconfig", "parse_config_h", "parse_config_h"),
    ("site", "USER_BASE", "USER_BASE"),
    ("site", "USER_SITE", "USER_SITE"),
    ("site", "PREFIXES", "PREFIXES"),
    ("site", "getuserbase", "getuserbase"),
    ("site", "getusersitepackages", "getusersitepackages"),
    ("site", "makepath", "makepath"),
    ("site", "getsitepackages", "getsitepackages"),
    ("site", "addpackage", "addpackage"),
    ("site", "addsitedir", "addsitedir"),
    ("site", "sethelper", "sethelper"),
    ("site", "setcopyright", "setcopyright"),
    ("site", "setquit", "setquit"),
    ("site", "main", "main"),
    ("traceback", "format_exc", "format_exc"),
    ("traceback", "format_exception", "format_exception"),
    (
        "traceback",
        "format_exception_only",
        "format_exception_only",
    ),
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
    ("tracemalloc", "start", "start"),
    ("tracemalloc", "stop", "stop"),
    ("tracemalloc", "is_tracing", "is_tracing"),
    ("tracemalloc", "take_snapshot", "take_snapshot"),
    ("tracemalloc", "get_traced_memory", "get_traced_memory"),
    ("tracemalloc", "get_tracemalloc_memory", "get_tracemalloc_memory"),
    ("tracemalloc", "reset_peak", "reset_peak"),
    ("tracemalloc", "get_object_traceback", "get_object_traceback"),
    ("tracemalloc", "Snapshot", "Snapshot"),
    ("tracemalloc", "Filter", "Filter"),
    ("types", "SimpleNamespace", "SimpleNamespace"),
    ("types", "MappingProxyType", "MappingProxyType"),
    ("types", "MethodType", "MethodType"),
    ("types", "FunctionType", "FunctionType"),
    ("types", "LambdaType", "LambdaType"),
    ("types", "GeneratorType", "GeneratorType"),
    ("types", "CoroutineType", "CoroutineType"),
    ("types", "DynamicClassAttribute", "DynamicClassAttribute"),
    ("types", "new_class", "new_class"),
    ("types", "resolve_bases", "resolve_bases"),
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
        &[
            "drive", "root", "anchor", "name", "stem", "suffix", "suffixes", "parts", "parent",
        ],
    ),
    (
        "PurePosixPath",
        &[
            "drive", "root", "anchor", "name", "stem", "suffix", "suffixes", "parts", "parent",
        ],
    ),
    (
        "PureWindowsPath",
        &[
            "drive", "root", "anchor", "name", "stem", "suffix", "suffixes", "parts", "parent",
        ],
    ),
    (
        "Path",
        &[
            "drive", "root", "anchor", "name", "stem", "suffix", "suffixes", "parts", "parent",
        ],
    ),
    ("ConfigParser", &["optionxform"]),
    ("RawConfigParser", &["optionxform"]),
];

pub const CLASS_WRITABLE_PROPERTIES: &[(&str, &[&str])] = &[
    ("ConfigParser", &["optionxform"]),
    ("RawConfigParser", &["optionxform"]),
];

pub const CLASS_MODULES: &[(&str, &str)] = &[
    ("JSONDecoder", "json.decoder"),
    ("JSONEncoder", "json.encoder"),
    ("SequenceMatcher", "difflib"),
    ("Differ", "difflib"),
    ("HtmlDiff", "difflib"),
    ("Mock", "unittest.mock"),
    ("MagicMock", "unittest.mock"),
    ("PropertyMock", "unittest.mock"),
];

/// Core classes whose instances accept arbitrary Python attributes.
///
/// User classes get this behaviour by default unless they declare slots; these
/// AST-declared classes bypass the source walk, so the walker needs the same
/// fact as metadata.
pub const DYNAMIC_ATTR_CLASSES: &[&str] = &["Namespace", "Mock", "MagicMock", "PropertyMock"];

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
            if *module == "contextvars" {
                out.push("Context");
            }
            if *module == "decimal" {
                out.extend(decimal::EXCEPTIONS.iter().map(|(name, _)| *name));
            }
            if *module == "graphlib" {
                out.extend(graphlib::EXCEPTIONS.iter().map(|(name, _)| *name));
            }
            if *module == "socket" {
                out.extend(socket::EXCEPTIONS.iter().map(|(name, _)| *name));
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
    ("ellipsis", &["__repr__"]),
    (
        "__py_SystemRandom",
        &["random", "randint", "getrandbits", "randbytes"],
    ),
    (
        "IPv4Address",
        &[
            "version",
            "_int",
            "_text",
            "compressed",
            "exploded",
            "packed",
            "is_private",
            "is_loopback",
            "is_multicast",
            "is_global",
        ],
    ),
    (
        "IPv4Network",
        &[
            "version",
            "prefixlen",
            "num_addresses",
            "_base",
            "network_address",
            "netmask",
            "hostmask",
            "broadcast_address",
        ],
    ),
    ("IPv4Interface", &["version", "prefixlen", "ip", "network"]),
    // ⛔ The DUNDERS are part of the surface, not decoration: `<` and `/` are
    // lowered against `class_has_attr`, so a class whose census names only its
    // fields compares and divides as a bare object — `Fraction(1,2) < …`
    // answered False while `.__lt__(…)` answered True.
    (
        "Decimal",
        &[
            "_c",
            "_e",
            "_kind",
            "__str__",
            "__repr__",
            "__add__",
            "__sub__",
            "__mul__",
            "__truediv__",
            "__floordiv__",
            "__mod__",
            "__pow__",
            "__neg__",
            "__pos__",
            "__abs__",
            "__float__",
            "__int__",
            "__hash__",
            "__eq__",
            "__ne__",
            "__lt__",
            "__le__",
            "__gt__",
            "__ge__",
            "is_nan",
            "is_infinite",
            "is_finite",
            "is_signed",
            "is_zero",
            "copy_abs",
            "copy_negate",
            "compare",
            "compare_total",
            "as_tuple",
            "sqrt",
            "quantize",
            "normalize",
            "to_integral_value",
        ],
    ),
    ("DecimalTuple", &["sign", "digits", "exponent"]),
    ("__PyDiffMatch", &["a", "b", "size"]),
    (
        "SequenceMatcher",
        &[
            "a",
            "b",
            "isjunk",
            "autojunk",
            "set_seq1",
            "set_seq2",
            "set_seqs",
            "ratio",
            "quick_ratio",
            "real_quick_ratio",
            "find_longest_match",
            "get_matching_blocks",
            "get_opcodes",
        ],
    ),
    ("Differ", &["compare"]),
    ("HtmlDiff", &["make_table", "make_file"]),
    ("__PyMockAny", &["__repr__", "__eq__"]),
    (
        "__PyMockCall",
        &["name", "args", "kwargs", "__repr__", "__str__", "__eq__"],
    ),
    ("__PyMockCallFactory", &["__call__", "__getattr__"]),
    ("__PyMockNamedCallFactory", &["name", "__call__"]),
    (
        "Mock",
        &[
            "return_value",
            "side_effect",
            "call_count",
            "called",
            "call_args",
            "call_args_list",
            "mock_calls",
            "_children",
            "_parent",
            "_parent_name",
            "_sealed",
            "_spec",
            "_spec_attrs",
            "__call__",
            "__mock_call__",
            "__getattr__",
            "__mock_assert_called_with__",
            "__mock_assert_called_once_with__",
            "assert_called_with",
            "assert_called_once_with",
            "assert_called_once",
            "assert_not_called",
            "assert_has_calls",
            "reset_mock",
            "attach_mock",
        ],
    ),
    (
        "MagicMock",
        &[
            "return_value",
            "side_effect",
            "call_count",
            "called",
            "call_args",
            "call_args_list",
            "mock_calls",
            "_children",
            "_parent",
            "_parent_name",
            "_sealed",
            "_spec",
            "_spec_attrs",
            "_mock_str",
            "_mock_len",
            "__call__",
            "__mock_call__",
            "__getattr__",
            "__str__",
            "__len__",
        ],
    ),
    (
        "PropertyMock",
        &[
            "return_value",
            "side_effect",
            "call_count",
            "called",
            "call_args",
            "call_args_list",
            "mock_calls",
            "_children",
            "_parent",
            "_parent_name",
            "_sealed",
            "_spec",
            "_spec_attrs",
            "__call__",
            "__mock_call__",
        ],
    ),
    ("__PyPatch", &["target", "name", "new", "old", "__enter__", "__exit__", "__call__"]),
    ("__PyPatchDict", &["target", "values", "clear", "old", "__enter__", "__exit__"]),
    ("__PyPatchFactory", &["__call__", "object", "dict"]),
    (
        "PurePath",
        &[
            "_s",
            "_win",
            "_is_win",
            "_make",
            "with_name",
            "with_suffix",
            "with_stem",
            "match",
            "joinpath",
            "__truediv__",
            "relative_to",
            "is_relative_to",
            "as_posix",
            "as_uri",
            "is_absolute",
            "is_reserved",
            "__eq__",
            "__hash__",
            "__str__",
            "__repr__",
        ],
    ),
    (
        "PurePosixPath",
        &[
            "_s",
            "_win",
            "_is_win",
            "_make",
            "with_name",
            "with_suffix",
            "with_stem",
            "match",
            "joinpath",
            "__truediv__",
            "relative_to",
            "is_relative_to",
            "as_posix",
            "as_uri",
            "is_absolute",
            "is_reserved",
            "__eq__",
            "__hash__",
            "__str__",
            "__repr__",
        ],
    ),
    (
        "PureWindowsPath",
        &[
            "_s",
            "_win",
            "_is_win",
            "_make",
            "with_name",
            "with_suffix",
            "with_stem",
            "match",
            "joinpath",
            "__truediv__",
            "relative_to",
            "is_relative_to",
            "as_posix",
            "as_uri",
            "is_absolute",
            "is_reserved",
            "__eq__",
            "__hash__",
            "__str__",
            "__repr__",
        ],
    ),
    (
        "Path",
        &[
            "_s",
            "_win",
            "_is_win",
            "_make",
            "with_name",
            "with_suffix",
            "with_stem",
            "match",
            "joinpath",
            "__truediv__",
            "relative_to",
            "is_relative_to",
            "as_posix",
            "as_uri",
            "is_absolute",
            "is_reserved",
            "__eq__",
            "__hash__",
            "__str__",
            "__repr__",
            "exists",
            "is_dir",
            "is_file",
            "stat",
            "lstat",
            "read_text",
            "read_bytes",
            "write_text",
            "write_bytes",
            "mkdir",
            "rmdir",
            "unlink",
            "iterdir",
            "glob",
            "rglob",
            "rename",
            "replace",
            "samefile",
            "resolve",
            "absolute",
            "expanduser",
            "touch",
            "chmod",
            "hardlink_to",
            "symlink_to",
            "open",
        ],
    ),
    (
        "StringIO",
        &[
            "_buf",
            "_pos",
            "closed",
            "write",
            "writelines",
            "getvalue",
            "read",
            "read1",
            "readline",
            "readlines",
            "__iter__",
            "seek",
            "tell",
            "truncate",
            "readable",
            "writable",
            "seekable",
            "flush",
            "detach",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
    (
        "BytesIO",
        &[
            "_buf",
            "_pos",
            "closed",
            "write",
            "writelines",
            "getvalue",
            "getbuffer",
            "read",
            "read1",
            "readline",
            "readlines",
            "__iter__",
            "seek",
            "tell",
            "truncate",
            "readable",
            "writable",
            "seekable",
            "flush",
            "detach",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
    (
        "Queue",
        &[
            "maxsize",
            "_items",
            "_unfinished_tasks",
            "put",
            "put_nowait",
            "get",
            "get_nowait",
            "empty",
            "full",
            "qsize",
            "task_done",
            "join",
        ],
    ),
    (
        "LifoQueue",
        &[
            "maxsize",
            "_items",
            "_unfinished_tasks",
            "put",
            "put_nowait",
            "get",
            "get_nowait",
            "empty",
            "full",
            "qsize",
            "task_done",
            "join",
        ],
    ),
    (
        "PriorityQueue",
        &[
            "maxsize",
            "_items",
            "_unfinished_tasks",
            "put",
            "put_nowait",
            "get",
            "get_nowait",
            "empty",
            "full",
            "qsize",
            "task_done",
            "join",
        ],
    ),
    (
        "SimpleQueue",
        &[
            "maxsize",
            "_items",
            "_unfinished_tasks",
            "put",
            "put_nowait",
            "get",
            "get_nowait",
            "empty",
            "full",
            "qsize",
            "task_done",
            "join",
        ],
    ),
    (
        "OptionParser",
        &["add_option", "add_option_group", "parse_args", "_options"],
    ),
    (
        "OptionGroup",
        &["add_option", "parser", "title", "description", "_options"],
    ),
    (
        "NullTranslations",
        &[
            "_info",
            "_charset",
            "_fallback",
            "add_fallback",
            "gettext",
            "ngettext",
            "pgettext",
            "npgettext",
            "info",
            "charset",
            "install",
        ],
    ),
    (
        "GNUTranslations",
        &[
            "_info",
            "_charset",
            "_fallback",
            "_catalog",
            "plural",
            "add_fallback",
            "gettext",
            "ngettext",
            "pgettext",
            "npgettext",
            "info",
            "charset",
            "install",
        ],
    ),
    (
        "TokenInfo",
        &["type", "string", "start", "end", "line", "exact_type", "__iter__"],
    ),
    (
        "__PyTraceback",
        &["_frames", "__len__", "__getitem__", "__iter__"],
    ),
    (
        "ConfigParser",
        &[
            "_sections",
            "_defaults",
            "_allow_no_value",
            "_raw",
            "_optionxform",
            "optionxform",
            "read_string",
            "read_dict",
            "sections",
            "has_section",
            "add_section",
            "set",
            "has_option",
            "remove_option",
            "remove_section",
            "options",
            "items",
            "get",
            "getint",
            "getfloat",
            "getboolean",
            "_normalize_option",
            "_merge_multiline_values",
            "_merge_allow_no_value",
            "_interpolate",
            "defaults",
            "__getitem__",
            "__setitem__",
            "write",
            "__contains__",
        ],
    ),
    (
        "RawConfigParser",
        &[
            "_sections",
            "_defaults",
            "_allow_no_value",
            "_raw",
            "_optionxform",
            "optionxform",
            "read_string",
            "read_dict",
            "sections",
            "has_section",
            "add_section",
            "set",
            "has_option",
            "remove_option",
            "remove_section",
            "options",
            "items",
            "get",
            "getint",
            "getfloat",
            "getboolean",
            "_normalize_option",
            "_merge_multiline_values",
            "_merge_allow_no_value",
            "_interpolate",
            "defaults",
            "__getitem__",
            "__setitem__",
            "write",
            "__contains__",
        ],
    ),
    (
        "VybeSocketImpl",
        &[
            "family",
            "sock_kind",
            "proto",
            "_timeout",
            "_opts",
            "_closed",
            "_rx",
            "_tx",
            "_listener",
            "_res",
            "settimeout",
            "gettimeout",
            "setblocking",
            "setsockopt",
            "getsockopt",
            "_addr_tuple",
            "bind",
            "listen",
            "getsockname",
            "getpeername",
            "fileno",
            "accept",
            "connect",
            "send",
            "sendall",
            "recv",
            "shutdown",
            "close",
            "dup",
            "detach",
            "makefile",
            "__enter__",
            "__exit__",
        ],
    ),
    (
        "HTTPStatus",
        &[
            "value",
            "phrase",
            "description",
            "is_informational",
            "is_success",
            "is_redirection",
            "is_client_error",
            "is_server_error",
            "__str__",
            "__repr__",
        ],
    ),
    (
        "HTTPMessage",
        &[
            "_headers",
            "get",
            "items",
            "keys",
            "__getitem__",
            "get_content_type",
        ],
    ),
    (
        "HTTPResponse",
        &[
            "status",
            "reason",
            "_sock",
            "_body_text",
            "_body_pos",
            "_closed",
            "version",
            "headers",
            "read",
            "readline",
            "readinto",
            "begin",
            "getheader",
            "getheaders",
            "isclosed",
            "close",
        ],
    ),
    (
        "HTTPConnection",
        &[
            "host",
            "port",
            "timeout",
            "sock",
            "_response",
            "_tunnel_host",
            "_tunnel_port",
            "connect",
            "putrequest",
            "putheader",
            "endheaders",
            "set_tunnel",
            "request",
            "getresponse",
            "close",
        ],
    ),
    (
        "Morsel",
        &[
            "key",
            "value",
            "coded_value",
            "expires",
            "path",
            "comment",
            "domain",
            "max-age",
            "secure",
            "version",
            "httponly",
            "samesite",
            "set",
            "__getitem__",
            "__setitem__",
            "__contains__",
            "keys",
            "OutputString",
            "__str__",
        ],
    ),
    (
        "SimpleCookie",
        &[
            "_cookies",
            "__setitem__",
            "__getitem__",
            "__len__",
            "keys",
            "clear",
            "load",
            "output",
            "js_output",
        ],
    ),
    ("CookieJar", &["_cookies", "__iter__"]),
    (
        "LWPCookieJar",
        &["_cookies", "filename", "__iter__", "save", "load"],
    ),
    ("Request", &["full_url", "data", "headers"]),
    ("HTMLParser", &["feed"]),
    (
        "__pprint_PrettyPrinter",
        &[
            "indent",
            "width",
            "depth",
            "stream",
            "compact",
            "sort_dicts",
            "underscore_numbers",
            "pformat",
            "pprint",
            "isrecursive",
            "isreadable",
        ],
    ),
    (
        "TopologicalSorter",
        &[
            "_preds",
            "_succs",
            "_prepared",
            "_ready",
            "_out",
            "_left",
            "_add_edges",
            "add",
            "_order",
            "_cycle_nodes",
            "prepare",
            "get_ready",
            "done",
            "is_active",
            "static_order",
        ],
    ),
    (
        "__PyNamedTempFile",
        &[
            "name",
            "mode",
            "delete",
            "closed",
            "_pos",
            "write",
            "writelines",
            "read",
            "readline",
            "readlines",
            "seek",
            "tell",
            "flush",
            "fileno",
            "readable",
            "writable",
            "seekable",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
    ("__PyDiskUsage", &["total", "used", "free"]),
    ("__PyTerminalSize", &["columns", "lines"]),
    ("Token", &["MISSING", "var", "old_value", "_used"]),
    (
        "ContextVar",
        &[
            "name",
            "_default",
            "_has_value",
            "_value",
            "get",
            "set",
            "reset",
        ],
    ),
    (
        "Context",
        &[
            "prec",
            "__enter__",
            "__exit__",
            "copy",
            "_items",
            "__len__",
            "items",
            "keys",
            "values",
            "get",
            "run",
        ],
    ),
    (
        "Fraction",
        &[
            "numerator",
            "denominator",
            "__str__",
            "__repr__",
            "__add__",
            "__radd__",
            "__sub__",
            "__mul__",
            "__rmul__",
            "__truediv__",
            "__floordiv__",
            "__mod__",
            "__pow__",
            "__neg__",
            "__pos__",
            "__abs__",
            "__float__",
            "__int__",
            "__hash__",
            "__eq__",
            "__ne__",
            "__lt__",
            "__le__",
            "__gt__",
            "__ge__",
            "limit_denominator",
        ],
    ),
    (
        "SchedEvent",
        &[
            "time",
            "priority",
            "action",
            "argument",
            "kwargs",
            "__lt__",
            "__le__",
            "__gt__",
            "__ge__",
            "__eq__",
        ],
    ),
    (
        "scheduler",
        &[
            "timefunc",
            "delayfunc",
            "_queue",
            "enter",
            "enterabs",
            "cancel",
            "empty",
            "queue",
            "run",
        ],
    ),
    ("__AsyncGenCM", &["gen", "__aenter__", "__aexit__"]),
    ("Namespace", &[]),
    (
        "ArgumentParser",
        &[
            "_options",
            "_positionals",
            "_defaults",
            "_mutex_groups",
            "_subparsers",
            "_subparser_dest",
            "add_argument",
            "add_argument_group",
            "add_mutually_exclusive_group",
            "add_subparsers",
            "set_defaults",
            "parse_args",
        ],
    ),
    (
        "__ArgparseGroup",
        &["parser", "required", "_members", "add_argument"],
    ),
    (
        "__ArgparseSubparsers",
        &["parser", "dest", "add_parser"],
    ),
    ("Module", &["__source", "__mode", "_nodes", "body"]),
    ("Expression", &["__source", "__mode", "_nodes", "body"]),
    ("Constant", &["value", "_nodes"]),
    ("Assign", &["value", "_nodes"]),
    ("Name", &["value", "_nodes"]),
    ("BinOp", &["value", "_nodes"]),
    ("Add", &["value", "_nodes"]),
    ("FunctionDef", &["value", "_nodes"]),
    ("Return", &["value", "_nodes"]),
    ("NodeVisitor", &["visit", "generic_visit"]),
    ("NodeTransformer", &["visit", "generic_visit"]),
    ("JSONDecoder", &["decode", "raw_decode"]),
    ("JSONEncoder", &["default", "encode"]),
    ("SelectorKey", &["fileobj", "fd", "events", "data"]),
    (
        "SelectSelector",
        &[
            "_map",
            "register",
            "unregister",
            "modify",
            "get_key",
            "get_map",
            "select",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
    (
        "EpollSelector",
        &[
            "_map",
            "register",
            "unregister",
            "modify",
            "get_key",
            "get_map",
            "select",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
    (
        "KqueueSelector",
        &[
            "_map",
            "register",
            "unregister",
            "modify",
            "get_key",
            "get_map",
            "select",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
    (
        "PollSelector",
        &[
            "_map",
            "register",
            "unregister",
            "modify",
            "get_key",
            "get_map",
            "select",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
    (
        "DevpollSelector",
        &[
            "_map",
            "register",
            "unregister",
            "modify",
            "get_key",
            "get_map",
            "select",
            "close",
            "__enter__",
            "__exit__",
        ],
    ),
];

/// Instance data attributes for AST-built core classes. User classes collect
/// these while walking `self.x = ...`; core classes bypass that walk, so the
/// same lowering metadata has to be seeded explicitly.
pub const CLASS_DATA_ATTRS: &[(&str, &[&str])] = &[(
    "TokenInfo",
    &["type", "string", "start", "end", "line", "exact_type"],
), (
    "Token",
    &["var", "old_value", "_used"],
), (
    "ContextVar",
    &["name", "_default", "_has_value", "_value"],
), (
    "Context",
    &["_items"],
), (
    "Queue",
    &["maxsize", "_items", "_unfinished_tasks"],
), (
    "LifoQueue",
    &["maxsize", "_items", "_unfinished_tasks"],
), (
    "PriorityQueue",
    &["maxsize", "_items", "_unfinished_tasks"],
), (
    "SimpleQueue",
    &["maxsize", "_items", "_unfinished_tasks"],
), (
    "__PyTraceFrame",
    &["filename", "lineno"],
), (
    "__PyTraceback",
    &["_frames"],
), (
    "__PyTraceStat",
    &["size", "count", "traceback"],
), (
    "__PyTraceStatDiff",
    &["size", "count", "size_diff", "count_diff", "traceback"],
), (
    "Snapshot",
    &["_stats"],
), (
    "Filter",
    &["inclusive", "filename_pattern", "lineno", "all_frames", "domain"],
), (
    "SchedEvent",
    &["time", "priority", "action", "argument", "kwargs"],
), (
    "scheduler",
    &["timefunc", "delayfunc", "_queue", "queue"],
), (
    "OptionParser",
    &["_options"],
), (
    "OptionGroup",
    &["parser", "title", "description", "_options"],
), (
    "NullTranslations",
    &["_info", "_charset", "_fallback"],
), (
    "GNUTranslations",
    &["_info", "_charset", "_fallback", "_catalog", "plural"],
), (
    "Logger",
    &["name", "level", "handlers"],
), (
    "Handler",
    &["level"],
), (
    "StreamHandler",
    &["level"],
), (
    "FileHandler",
    &["level"],
), (
    "__WarningRecord",
    &["message", "category"],
), (
    "__CatchWarnings",
    &["record", "entries"],
), (
    "__py_shlex_class",
    &["text", "posix", "whitespace_split", "commenters"],
), (
    "__py_TextWrapper",
    &[
        "width",
        "initial_indent",
        "subsequent_indent",
        "break_long_words",
        "break_on_hyphens",
        "expand_tabs",
        "replace_whitespace",
        "drop_whitespace",
        "max_lines",
        "placeholder",
    ],
), (
    "__string_Template",
    &[
        "delimiter",
        "template",
        "substitute",
        "safe_substitute",
        "get_identifiers",
        "is_valid",
    ],
), (
    "__string_Formatter",
    &["format", "vformat"],
), (
    "__AsyncGenCM",
    &["gen"],
), (
    "ArgumentParser",
    &[
        "_options",
        "_positionals",
        "_defaults",
        "_mutex_groups",
        "_subparsers",
        "_subparser_dest",
    ],
), (
    "__ArgparseGroup",
    &["parser", "required", "_members"],
), (
    "__ArgparseSubparsers",
    &["parser", "dest"],
), (
    "Module",
    &["__source", "__mode", "_nodes", "body"],
), (
    "Expression",
    &["__source", "__mode", "_nodes", "body"],
), (
    "Constant",
    &["value", "_nodes"],
), (
    "Assign",
    &["value", "_nodes"],
), (
    "Name",
    &["value", "_nodes"],
), (
    "BinOp",
    &["value", "_nodes"],
), (
    "Add",
    &["value", "_nodes"],
), (
    "FunctionDef",
    &["value", "_nodes"],
), (
    "Return",
    &["value", "_nodes"],
), (
    "SelectorKey",
    &["fileobj", "fd", "events", "data"],
), (
    "SelectSelector",
    &["_map"],
), (
    "EpollSelector",
    &["_map"],
), (
    "KqueueSelector",
    &["_map"],
), (
    "PollSelector",
    &["_map"],
), (
    "DevpollSelector",
    &["_map"],
)];

/// The bytes helper surface. Not a module: the gate is `source_uses_bytes`,
/// the same one the prelude carried, so the walker asks for it directly rather
/// than through `MODULE_CLASSES`.
/// The `collections` helpers, gated by the NAME that reaches each one — the
/// walker rewrites `Counter(...)`, `deque(...)` and `defaultdict(...)` into
/// them, and the source always spells the class it asked for.
///
/// ⛔ The DEQUE helpers are unconditional: the walker emits
/// `__py_deque_append` for list appends generally, so `UserList([1]).append(2)`
/// reaches it without the word `deque` ever appearing.
fn collections_functions() -> Vec<Statement> {
    let mut out = vec![
        collections::user_dict(),
        collections::user_list(),
        collections::user_string(),
    ];
    out.extend(collections::deque_functions());
    out.extend(collections::chainmap_functions());
    out.extend(collections::ordereddict_functions());
    out.push(collections::ordereddict_front());
    out.extend(collections::defaultdict_functions());
    out.extend(collections::user_functions());
    out
}

pub const BYTES_CLASSES: &[&str] = &["__PyIncrementalEncoder", "__PyIncrementalDecoder"];

pub fn bytes_declarations() -> Vec<Statement> {
    let mut out = vec![bytes::incremental_encoder(), bytes::incremental_decoder()];
    out.extend(bytes::module_functions());
    out
}

/// `object()` — a bare instance with no attributes of its own, existing only
/// for identity (`a is a`, `a is not b`). Not a module either: the builtin is
/// always in scope, so the gate is the literal call text.
pub const OBJECT_CLASSES: &[&str] = &["__PyObject"];

pub fn object_declarations() -> Vec<Statement> {
    vec![object_class::bare_object()]
}

pub fn ellipsis_declarations() -> Vec<Statement> {
    vec![builtins::ellipsis_type(), builtins::ellipsis_binding()]
}

pub fn slice_declarations() -> Vec<Statement> {
    builtins::slice_declarations()
}

pub fn module_dunder_declarations() -> Vec<Statement> {
    builtins::module_dunder_declarations()
}

pub fn exception_group_declarations() -> Vec<Statement> {
    exceptions::exception_group_declarations()
}

/// The global a `<module>.<name>` read denotes, if this module declares it.
pub fn module_member(module: &str, name: &str) -> Option<&'static str> {
    MODULE_SURFACE
        .iter()
        .find(|(m, n, _)| *m == module && *n == name)
        .map(|(_, _, global)| *global)
}

pub fn class_module(name: &str) -> Option<&'static str> {
    CLASS_MODULES
        .iter()
        .find(|(class_name, _)| *class_name == name)
        .map(|(_, module)| *module)
}

pub fn core_class_parents(name: &str) -> Vec<&'static str> {
    let mut out = Vec::new();
    match name {
        "NodeTransformer" => out.push("NodeVisitor"),
        "EpollSelector" | "KqueueSelector" | "PollSelector" | "DevpollSelector" => {
            out.push("SelectSelector");
        }
        _ => {}
    }
    if name == "GNUTranslations" {
        out.push("NullTranslations");
    }
    for (child, parent) in warnings::CATEGORIES
        .iter()
        .chain(decimal::EXCEPTIONS.iter())
        .chain(graphlib::EXCEPTIONS.iter())
        .chain(socket::EXCEPTIONS.iter())
        .chain(queue::EXCEPTIONS.iter())
        .chain(http_ssl::EXCEPTIONS.iter())
    {
        if *child == name {
            out.push(*parent);
        }
    }
    out
}

/// Which classes a module's import needs. A class is not reachable by its own
/// name in Python — `ipaddress.IPv4Address` is, but a program far more often
/// only ever names `ip_address` — so the gate is the MODULE, not the class.
const MODULE_CLASSES: &[(&str, &[&str])] = &[
    (
        "ipaddress",
        &["IPv4Address", "IPv4Network", "IPv4Interface"],
    ),
    ("warnings", &["__WarningRecord", "__CatchWarnings"]),
    (
        "logging",
        &[
            "LogRecord",
            "Formatter",
            "Filter",
            "Handler",
            "StreamHandler",
            "FileHandler",
            "Logger",
        ],
    ),
    (
        "contextlib",
        &[
            "__NullContext",
            "__Closing",
            "__Suppress",
            "__GenCM",
            "__AsyncGenCM",
            "__Redirect",
        ],
    ),
    (
        "traceback",
        &["FrameSummary", "StackSummary", "TracebackException"],
    ),
    (
        "http",
        &[
            "HTTPMessage",
            "HTTPResponse",
            "HTTPConnection",
            "HTTPSConnection",
            "HTTPStatus",
            "Morsel",
            "SimpleCookie",
            "CookieJar",
            "LWPCookieJar",
        ],
    ),
    (
        "http.client",
        &[
            "HTTPMessage",
            "HTTPResponse",
            "HTTPConnection",
            "HTTPSConnection",
            "HTTPException",
            "BadStatusLine",
            "IncompleteRead",
            "CannotSendRequest",
        ],
    ),
    ("http.cookies", &["Morsel", "SimpleCookie", "CookieError"]),
    ("http.cookiejar", &["CookieJar", "LWPCookieJar"]),
    ("urllib.request", &["Request"]),
    ("html.parser", &["HTMLParser"]),
    ("ssl", &["SSLContext", "TLSVersion", "Purpose"]),
    (
        "csv",
        &[
            "__PyCsvExcel",
            "__PyCsvExcelTab",
            "__PyCsvSemi",
            "__PyCsvReader",
            "__PyCsvWriter",
            "__PyCsvDictReader",
            "__PyCsvDictWriter",
            "Sniffer",
        ],
    ),
    (
        "threading",
        &[
            "__PyLock",
            "Semaphore",
            "BoundedSemaphore",
            "Event",
            "Condition",
            "Barrier",
            "local",
            "Thread",
            "Timer",
        ],
    ),
    (
        "queue",
        &["Queue", "LifoQueue", "PriorityQueue", "SimpleQueue"],
    ),
    ("concurrent", &["Future"]),
    (
        "multiprocessing",
        &[
            "Process",
            "Pool",
            "__PyValue",
            "__PyProcessInfo",
            "Manager",
            "__PyPipeEnd",
        ],
    ),
    (
        "subprocess",
        &[
            "CompletedProcess",
            "CalledProcessError",
            "TimeoutExpired",
            "Popen",
        ],
    ),
    // No classes — the entry exists so `time`'s module FUNCTIONS splice.
    ("time", &[]),
    ("timeit", &["__TimeitTimer"]),
    ("tokenize", &["TokenInfo", "TokenError"]),
    (
        "tracemalloc",
        &[
            "__PyTraceFrame",
            "__PyTraceback",
            "__PyTraceStat",
            "__PyTraceStatDiff",
            "Snapshot",
            "Filter",
        ],
    ),
    ("pathlib", &["PurePath", "Path"]),
    ("fractions", &["Fraction"]),
    ("decimal", &["Context", "DecimalTuple", "Decimal"]),
    ("filecmp", &["__PyDirCmp"]),
    (
        "difflib",
        &["__PyDiffMatch", "SequenceMatcher", "Differ", "HtmlDiff"],
    ),
    (
        "unittest.mock",
        &[
            "__PyMockAny",
            "__PyMockCall",
            "__PyMockCallFactory",
            "__PyMockNamedCallFactory",
            "Mock",
            "MagicMock",
            "PropertyMock",
            "__PyPatch",
            "__PyPatchDict",
            "__PyPatchFactory",
        ],
    ),
    // No classes — the entry exists so `linecache`'s module FUNCTIONS splice.
    ("linecache", &[]),
    ("gettext", &["NullTranslations", "GNUTranslations"]),
    (
        "optparse",
        &["__PyOptValues", "OptionParser", "OptionGroup"],
    ),
    (
        "argparse",
        &[
            "Namespace",
            "ArgumentParser",
            "__ArgparseGroup",
            "__ArgparseSubparsers",
        ],
    ),
    (
        "ast",
        &[
            "Module",
            "Expression",
            "Constant",
            "Assign",
            "Name",
            "BinOp",
            "Add",
            "FunctionDef",
            "Return",
            "NodeVisitor",
            "NodeTransformer",
        ],
    ),
    ("plistlib", &[]),
    (
        "selectors",
        &[
            "SelectorKey",
            "SelectSelector",
            "EpollSelector",
            "KqueueSelector",
            "PollSelector",
            "DevpollSelector",
        ],
    ),
    ("json", &["JSONDecoder", "JSONEncoder"]),
    ("sched", &["SchedEvent", "scheduler"]),
    ("contextvars", &["Token", "ContextVar"]),
    ("tomllib", &["TOMLDecodeError"]),
    ("shutil", &["__PyDiskUsage", "__PyTerminalSize"]),
    // No classes — the entry exists so `site`'s module FUNCTIONS splice.
    ("site", &[]),
    ("io", &["StringIO", "BytesIO"]),
    // ⛔ NO class names. Seeding `UserList` into `py_defined_classes` before
    // the walk turns `UserList([1])` into a `New`, which fights the walker's
    // own `UserList` -> `__py_userlist` rewrite. The prelude declared the
    // classes without registering them; so does this.
    ("collections", &[]),
    (
        "configparser",
        &[
            "ConfigParser",
            "RawConfigParser",
            "BasicInterpolation",
            "ExtendedInterpolation",
            "Error",
            "DuplicateSectionError",
            "NoSectionError",
            "NoOptionError",
        ],
    ),
    ("socket", &["VybeSocketImpl"]),
    ("pprint", &["__pprint_PrettyPrinter"]),
    ("graphlib", &["TopologicalSorter"]),
    ("tempfile", &["__PyNamedTempFile"]),
    // No classes — the entry exists so `struct`'s module FUNCTIONS splice.
    ("struct", &[]),
    // No classes — the entry exists so `sysconfig`'s module FUNCTIONS splice.
    ("sysconfig", &[]),
    // No classes — the entry exists so `fnmatch`'s module FUNCTIONS splice.
    ("fnmatch", &[]),
    ("shlex", &["__py_shlex_class"]),
    ("textwrap", &["__py_TextWrapper"]),
    ("string", &["__string_Template", "__string_Formatter"]),
    ("random", &["__py_SystemRandom"]),
    // No classes — these are module-level values/functions only.
    ("types", &[]),
    // Not a module — the gate is the ANNOTATION machinery that builds it.
    ("__annotations__", &["__py_type_obj"]),
    // `typing.List[int]` and friends normalize to a GenericAlias carrying
    // runtime type objects.
    ("typing", &["__py_type_obj"]),
    ("__py_type_obj", &["__py_type_obj"]),
    ("get_annotations", &["__py_type_obj"]),
    // `inspect.signature()` materializes annotation objects even when the
    // program never reads `f.__annotations__` directly.
    ("signature", &["__py_type_obj"]),
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
        if *module == "contextvars" {
            for (owner, build) in MODULE_FUNCTIONS {
                if owner == module {
                    out.extend(build());
                }
            }
        }
        out.extend(generated_classes(module));
        for (name, build) in CORE_CLASSES {
            if classes.contains(name) && !is_user_declared(name) {
                out.push(build());
            }
        }
        if *module != "contextvars" {
            for (owner, build) in MODULE_FUNCTIONS {
                if owner == module {
                    out.extend(build());
                }
            }
        }
    }
    out
}
