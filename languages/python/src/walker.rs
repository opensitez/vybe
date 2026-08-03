use super::{PythonParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use std::collections::HashMap;
use vybe_ast::*;

// ════════════════════════════════════════════════════════════════════════════
// Indentation preprocessor
// ════════════════════════════════════════════════════════════════════════════
// Python uses indentation for blocks. pest cannot track indent state, so we
// insert explicit markers before parsing:
//   ⇥ (U+21E5) = INDENT
//   ⇤ (U+21E4) = DEDENT

/// Update bracket depth for one physical line, skipping string literals.
/// `in_triple` carries an open triple-quoted string across lines (holding its
/// quote char while one is open).
///
/// Brackets and `#` inside a string literal are content, not syntax. Counting
/// them makes logical-line resolution glue unrelated statements together —
/// `s = "{"` would swallow every following line up to the next `}`.
fn scan_line_brackets(line: &str, bracket_depth: &mut i32, in_triple: &mut Option<char>) {
    let b: Vec<char> = line.chars().collect();
    let mut i = 0;
    while i < b.len() {
        let c = b[i];
        if let Some(q) = *in_triple {
            if c == '\\' {
                i += 2;
                continue;
            }
            if c == q && b.get(i + 1) == Some(&q) && b.get(i + 2) == Some(&q) {
                *in_triple = None;
                i += 3;
                continue;
            }
            i += 1;
            continue;
        }
        match c {
            // Rest of the line is a comment.
            '#' => return,
            '(' | '[' | '{' => *bracket_depth += 1,
            ')' | ']' | '}' => *bracket_depth -= 1,
            '\'' | '"' => {
                if b.get(i + 1) == Some(&c) && b.get(i + 2) == Some(&c) {
                    *in_triple = Some(c);
                    i += 3;
                    continue;
                }
                // Single-quoted: consume through the closing quote. An
                // unterminated one just ends at EOL; the grammar reports it.
                i += 1;
                while i < b.len() {
                    if b[i] == '\\' {
                        i += 2;
                        continue;
                    }
                    if b[i] == c {
                        break;
                    }
                    i += 1;
                }
            }
            _ => {}
        }
        i += 1;
    }
}

fn preprocess_indentation(source: &str) -> String {
    // Phase 1: Resolve physical lines into logical lines.
    // Handles explicit continuation (backslash) and implicit continuation (unclosed brackets).
    let mut logical_lines: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut bracket_depth: i32 = 0;
    let mut in_triple: Option<char> = None;

    for line in source.lines() {
        let trimmed = line.trim();

        // During continuation, skip blank/comment lines — but never inside a
        // triple-quoted string, where both are literal content.
        if in_triple.is_none()
            && !current.is_empty()
            && (trimmed.is_empty() || trimmed.starts_with('#'))
        {
            continue;
        }

        let was_in_triple = in_triple.is_some();

        if current.is_empty() {
            current.push_str(line);
        } else if was_in_triple {
            // Inside a triple-quoted string the physical newline and the
            // leading whitespace are literal content — joining with a space
            // would rewrite the value (`print(f"""a\nb""")`).
            current.push('\n');
            current.push_str(line);
        } else {
            // Continuation: join with space + trimmed content
            current.push(' ');
            current.push_str(trimmed);
        }

        scan_line_brackets(line, &mut bracket_depth, &mut in_triple);

        // Explicit continuation: backslash at end of line (string content is
        // not a continuation marker).
        if !was_in_triple && in_triple.is_none() && line.trim_end().ends_with('\\') {
            if let Some(pos) = current.rfind('\\') {
                current.truncate(pos);
            }
            continue;
        }

        // Implicit continuation: unclosed brackets
        if bracket_depth > 0 {
            continue;
        }

        logical_lines.push(std::mem::take(&mut current));
        bracket_depth = 0;
    }
    if !current.is_empty() {
        logical_lines.push(current);
    }

    // Phase 2: Process indentation on logical lines
    let mut result = String::with_capacity(source.len() * 2);
    let mut indent_stack: Vec<usize> = vec![0];
    let mut first = true;

    for line in &logical_lines {
        // Count leading spaces (expand tabs to 8)
        let mut indent = 0;
        let mut chars = line.chars().peekable();
        while let Some(&c) = chars.peek() {
            match c {
                ' ' => {
                    indent += 1;
                    chars.next();
                }
                '\t' => {
                    indent += 8 - (indent % 8);
                    chars.next();
                }
                _ => break }
        }

        let rest: String = chars.collect();

        // Skip blank lines and comment-only lines
        let trimmed = rest.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            if !first {
                result.push('\n');
            }
            result.push_str(line);
            first = false;
            continue;
        }

        if !first {
            result.push('\n');
        }
        first = false;

        let current_indent = *indent_stack.last().unwrap();

        if indent > current_indent {
            indent_stack.push(indent);
            result.push('\u{21E5}'); // INDENT
        } else {
            while indent < *indent_stack.last().unwrap() {
                indent_stack.pop();
                result.push('\u{21E4}'); // DEDENT
            }
            // After popping, if indent is above the new top, it's a new block level
            if indent > *indent_stack.last().unwrap() {
                indent_stack.push(indent);
                result.push('\u{21E5}'); // INDENT
            }
        }

        result.push_str(line);
    }

    // Close remaining indents at EOF
    while indent_stack.len() > 1 {
        indent_stack.pop();
        result.push('\n');
        result.push('\u{21E4}');
    }

    result
}

// ════════════════════════════════════════════════════════════════════════════
// Entry point
// ════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    PY_IMPORTED_MODULES.with(|m| m.borrow_mut().clear());
    PY_FROM_IMPORTED_MODULES.with(|m| m.borrow_mut().clear());
    PY_FLOAT_RETURNING_IMPORTS.with(|m| m.borrow_mut().clear());
    PY_SYS_MODULES_BOUND.with(|b| b.set(false));
    PY_DYNAMIC_MODULE_VARS.with(|m| m.borrow_mut().clear());
    PY_DYNAMIC_MODULE_REGISTRY.with(|m| m.borrow_mut().clear());
    PY_DYNAMIC_MODULE_ATTRS.with(|m| m.borrow_mut().clear());
    PY_DYNAMIC_MODULE_ALL.with(|m| m.borrow_mut().clear());
    PY_STRING_CONSTS.with(|m| m.borrow_mut().clear());
    PY_MIMETYPE_CUSTOMS.with(|m| m.borrow_mut().clear());
    PY_NONE_VARS.with(|m| m.borrow_mut().clear());
    PY_MAPPING_PROXY_VARS.with(|m| m.borrow_mut().clear());
    PY_SIMPLE_NAMESPACE_VARS.with(|m| m.borrow_mut().clear());
    PY_DEFINED_CLASSES.with(|m| m.borrow_mut().clear());
    PY_DEFINED_FUNCTIONS.with(|m| m.borrow_mut().clear());
    PY_CALLABLE_CLASSES.with(|m| m.borrow_mut().clear());
    PY_CLASSES_WITH_INIT.with(|m| m.borrow_mut().clear());
    PY_CLASS_PARENTS.with(|m| m.borrow_mut().clear());
    PY_CLASS_ATTRS.with(|m| m.borrow_mut().clear());
    PY_CLASS_DATA_ATTRS.with(|m| m.borrow_mut().clear());
    PY_INSTANCE_CLASSES.with(|m| m.borrow_mut().clear());
    PY_INSTANCE_ATTRS.with(|m| m.borrow_mut().clear());
    PY_ASSIGN_TARGET_DEPTH.with(|d| d.set(0));
    PY_NAMEDTUPLE_DEFS.with(|m| m.borrow_mut().clear());
    PY_NAMEDTUPLE_INSTANCES.with(|m| m.borrow_mut().clear());
    PY_SQL_VARS.with(|m| m.borrow_mut().clear());
    PY_RE_VARS.with(|m| m.borrow_mut().clear());
    PY_COUNTER_VARS.with(|m| m.borrow_mut().clear());
    PY_DEFAULTDICT_VARS.with(|m| m.borrow_mut().clear());
    PY_DEQUE_MAXLEN_VARS.with(|m| m.borrow_mut().clear());
    PY_CHAINMAP_VARS.with(|m| m.borrow_mut().clear());
    PY_ITERATOR_VARS.with(|m| m.borrow_mut().clear());
    PY_GENERATOR_FUNCS.with(|m| m.borrow_mut().clear());
    PY_GENERATOR_VARS.with(|m| m.borrow_mut().clear());
    PY_USERLIST_VARS.with(|m| m.borrow_mut().clear());
    PY_USERDICT_VARS.with(|m| m.borrow_mut().clear());
    PY_DICT_VARS.with(|m| m.borrow_mut().clear());
    PY_SET_VARS.with(|m| m.borrow_mut().clear());
    let preprocessed = preprocess_indentation(source);
    let pairs = PythonParser::parse(Rule::program, &preprocessed)
        .map_err(|e| format!("Parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();

    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => {
                walk_stmt_into(top, &mut body, &mut imports)?;
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE => continue,
                _ => walk_stmt_into(pair, &mut body, &mut imports)? }
        }
    }

    apply_float_var_repr(&mut body, &mut HashMap::new());

    // Prepend the bytes-repr source helper when the program uses bytes, so
    // `b'…'` display resolves to a real `__vybe_bytes_repr` function.
    if source_uses_bytes(source) {
        let mut prelude = parse_python_prelude(BYTES_REPR_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // Define the `Ellipsis` singleton (bound to `...`) as a real object so it is
    // distinct from `None`, is its own singleton (`... is ...`), and reprs as
    // `Ellipsis` with `type(...).__name__ == "ellipsis"`.
    if source_uses_ellipsis(source) {
        let mut prelude = parse_python_prelude(ELLIPSIS_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // Define the `slice` type when `slice(...)` is constructed, so each call
    // yields a fresh object (`slice(1) is slice(1)` → False).
    if source_uses_slice_ctor(source) {
        let mut prelude = parse_python_prelude(SLICE_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    if source.contains("ExceptionGroup") {
        let mut prelude = parse_python_prelude(EXCEPTION_GROUP_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // Pure-Python `logging` module surface (getLogger/Logger + handler/formatter
    // classes). Constants come from [namespace_constants].
    if source.contains("import logging") {
        let mut prelude = parse_python_prelude(LOGGING_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // Pure-Python `warnings` module: warning classes + a recording
    // `catch_warnings` context manager; the filters are no-ops.
    if source.contains("import warnings") {
        let mut prelude = parse_python_prelude(WARNINGS_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // Pure-Python `traceback` module (best-effort formatting; content needing the
    // live exception is a header stub since `sys.exc_info()` isn't populated).
    if source.contains("import traceback") {
        let mut prelude = parse_python_prelude(TRACEBACK_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `contextlib` context-manager helpers as global names (covers
    // `from contextlib import nullcontext/closing/suppress`).
    if source.contains("contextlib") {
        let mut prelude = parse_python_prelude(CONTEXTLIB_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `http.client` / `ssl` — surface, not transport (see the prelude note).
    // Gated on either import; both share one prelude because `HTTPSConnection`
    // and `ssl.wrap_socket` refer to each other.
    if source.contains("import http") || source.contains("import ssl") {
        let mut prelude = parse_python_prelude(HTTP_SSL_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `ipaddress` — pure address arithmetic; a prelude because the addresses
    // and networks are classes.
    if source.contains("import ipaddress") {
        let mut prelude = parse_python_prelude(IPADDRESS_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `socket` — the class is genuinely stateful (WASI resource + streams +
    // timeout + option table), which is the one case a prelude is for.
    if source.contains("import socket") {
        let mut prelude = parse_python_prelude(SOCKET_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `io.StringIO` — an in-memory text stream (pure Python, no host I/O), so
    // `print(..., file=buf)` and manual read/write/seek work against a buffer.
    if source.contains("import io") {
        let mut prelude = parse_python_prelude(IO_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `fnmatch` — shell-style filename matching (pure Python, self-contained
    // iterative `*`/`?` matcher; no `re` dependency).
    if source.contains("import fnmatch") {
        let mut prelude = parse_python_prelude(FNMATCH_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `configparser` — INI parsing (pure Python; sections stored as nested
    // dicts so no string data lives in a `self.attr` slice/concat).
    if source.contains("import configparser") {
        let mut prelude = parse_python_prelude(CONFIGPARSER_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `pathlib` — PurePath/Path (pure Python; string-only path math + FS-method
    // stubs). Gated on the bare substring so `from pathlib import PurePath/Path`
    // and friends all trigger injection, not just `import pathlib`.
    if source.contains("pathlib") {
        let mut prelude = parse_python_prelude(PATHLIB_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `os.path` — pure-string POSIX helpers (join/split/normpath/…) are emitter
    // adapters (common:python.ospath_*), emitted at the call site. No prelude.

    // `random` — distributions + range/weight helpers over `random.random()`
    // + `math` (no host RNG beyond the base entropy source).
    if source.contains("random") {
        let mut prelude = parse_python_prelude(RANDOM_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }

    // `string` module classes/functions (Template/Formatter/capwords). Constants
    // are intercepted at the member-read site; the class/function surface needs
    // real definitions, so inject them when the module is imported.
    if source.contains("import string") {
        let mut prelude = parse_python_prelude(STRING_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("shlex") {
        let mut prelude = parse_python_prelude(SHLEX_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("textwrap") {
        let mut prelude = parse_python_prelude(TEXTWRAP_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("pprint") {
        let mut prelude = parse_python_prelude(PPRINT_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("functools") {
        let mut prelude = parse_python_prelude(FUNCTOOLS_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("__annotations__") {
        let mut prelude = parse_python_prelude(TYPEOBJ_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("|=") {
        let mut prelude = parse_python_prelude(DICT_OP_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("__iadd__") {
        let mut prelude = parse_python_prelude(LIST_IADD_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("collections") {
        let mut prelude = parse_collections_prelude();
        prelude.append(&mut body);
        body = prelude;
    }
    if source.contains("types") {
        let mut prelude = parse_python_prelude(TYPES_PRELUDE);
        prelude.append(&mut body);
        body = prelude;
    }
    body.insert(
        0,
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident("__doc__")],
            value: Expression::new(ExprKind::Lit(Literal::Null)), by_ref: false }),
    );
    Ok(Module {
        name: "main".into(),
        language: Lang::Python,
        body,
        imports })
}

/// Heuristic: does the source reference bytes at all? Only gates whether the
/// repr helper is injected — a false positive just adds an unused function.
fn source_uses_bytes(source: &str) -> bool {
    source.contains("b'")
        || source.contains("b\"")
        || source.contains("B'")
        || source.contains("B\"")
        || source.contains("bytes")
        || source.contains("bytearray")
        || source.contains(".hex(")
        || source.contains(".encode(")
        || source.contains(".decode(")
}

/// Parse a Python source prelude into top-level statements. Errors yield `[]`
/// so a prelude problem can never break user compilation.
///
/// Memoised per PROCESS, the way `vybe_language_js::prelude_body` is. The
/// preludes are conditional (only what the source references is injected) but
/// each one still went back through pest on every compile, and there are 27
/// injection sites. Cloning a parsed body is far cheaper than re-parsing it.
///
/// Keyed by CONTENT rather than pointer: `parse_collections_prelude` builds its
/// source at run time from the counters it saw, so a pointer key would be both
/// wrong and unsound once that string is freed.
fn parse_python_prelude(src: &str) -> Vec<Statement> {
    use std::collections::HashMap;
    use std::sync::{Mutex, OnceLock};
    static CACHE: OnceLock<Mutex<HashMap<String, Vec<Statement>>>> = OnceLock::new();
    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    if let Some(hit) = cache.lock().unwrap().get(src) {
        return hit.clone();
    }
    let parsed = parse_python_prelude_uncached(src);
    // A dynamically-built prelude varies per program; caching every variant
    // would grow without bound in a long-lived worker.
    if cache.lock().unwrap().len() < 256 {
        cache.lock().unwrap().insert(src.to_string(), parsed.clone());
    }
    parsed
}

fn parse_python_prelude_uncached(src: &str) -> Vec<Statement> {
    let preprocessed = preprocess_indentation(src);
    let pairs = match PythonParser::parse(Rule::program, &preprocessed) {
        Ok(p) => p,
        Err(e) => {
            if std::env::var("VYBE_PRELUDE_DEBUG").is_ok() {
                eprintln!("[prelude-parse-error] {e}");
                for (i, l) in preprocessed.lines().enumerate() {
                    eprintln!("[pp {:3}] {l}", i + 1);
                }
            }
            return Vec::new();
        }
    };
    let mut body = Vec::new();
    let mut imports = Vec::new();
    for top in pairs {
        match top.as_rule() {
            Rule::program => {
                for pair in top.into_inner() {
                    match pair.as_rule() {
                        Rule::EOI | Rule::NEWLINE => continue,
                        _ => {
                            let _ = walk_stmt_into(pair, &mut body, &mut imports);
                        }
                    }
                }
            }
            Rule::EOI => continue,
            _ => {
                let _ = walk_stmt_into(top, &mut body, &mut imports);
            }
        }
    }
    body
}

fn parse_collections_prelude() -> Vec<Statement> {
    let counters = PY_COUNTER_VARS.with(|m| {
        let snapshot = m.borrow().clone();
        m.borrow_mut().clear();
        snapshot
    });
    let defaultdicts = PY_DEFAULTDICT_VARS.with(|m| {
        let snapshot = m.borrow().clone();
        m.borrow_mut().clear();
        snapshot
    });
    let deques = PY_DEQUE_MAXLEN_VARS.with(|m| {
        let snapshot = m.borrow().clone();
        m.borrow_mut().clear();
        snapshot
    });
    let chainmaps = PY_CHAINMAP_VARS.with(|m| {
        let snapshot = m.borrow().clone();
        m.borrow_mut().clear();
        snapshot
    });
    let body = parse_python_prelude(COLLECTIONS_PRELUDE);
    PY_COUNTER_VARS.with(|m| *m.borrow_mut() = counters);
    PY_DEFAULTDICT_VARS.with(|m| *m.borrow_mut() = defaultdicts);
    PY_DEQUE_MAXLEN_VARS.with(|m| *m.borrow_mut() = deques);
    PY_CHAINMAP_VARS.with(|m| *m.borrow_mut() = chainmaps);
    body
}

/// Python source for `__vybe_bytes_repr(int_array) -> "b'…'"`. Escape fragments
/// are built from `chr(92)` (backslash) rather than backslash string literals,
/// which the Python string-escape lowering mishandles.
/// Does the source reference `...` or `Ellipsis`? Gates the singleton prelude
/// (a false positive just adds an unused class + binding).
fn source_uses_ellipsis(source: &str) -> bool {
    source.contains("...") || source.contains("Ellipsis")
}

/// Does the source call the `slice()` constructor (as opposed to `.slice(` method
/// calls or `islice`)? Requires `slice(` whose preceding char is not part of an
/// identifier and not a `.` member access. A false positive only adds an unused
/// class definition.
fn source_uses_slice_ctor(source: &str) -> bool {
    let bytes = source.as_bytes();
    let mut i = 0;
    while let Some(pos) = source[i..].find("slice(") {
        let at = i + pos;
        let prev_ok = at == 0
            || !matches!(bytes[at - 1], b'.' | b'_') && !bytes[at - 1].is_ascii_alphanumeric();
        if prev_ok {
            return true;
        }
        i = at + 5;
    }
    false
}

/// The `Ellipsis` singleton — a real object (so `... is None` is False and
/// `type(...).__name__` is `"ellipsis"`), bound once at module scope.
const ELLIPSIS_PRELUDE: &str = r#"
class ellipsis:
    def __repr__(self):
        return "Ellipsis"
Ellipsis = ellipsis()
"#;

/// The `slice` type — a real object per construction, so `slice(1) is slice(1)`
/// is False and `.start`/`.stop`/`.step` are readable. `slice(stop)`,
/// `slice(start, stop)`, and `slice(start, stop, step)` follow CPython's arity
/// rules.
const SLICE_PRELUDE: &str = r#"
__slice_unset = object()
class slice:
    def __init__(self, start, stop=__slice_unset, step=None):
        if stop is __slice_unset:
            self.start = None
            self.stop = start
            self.step = None
        else:
            self.start = start
            self.stop = stop
            self.step = step
    def __repr__(self):
        a = repr(self.start)
        b = repr(self.stop)
        c = repr(self.step)
        return "slice(" + a + ", " + b + ", " + c + ")"
"#;

const EXCEPTION_GROUP_PRELUDE: &str = r#"
def __py_exception_group_split(eg, typ):
    matched = []
    rest = []
    for e in eg.exceptions:
        if __py_type_name(e) == typ:
            matched.append(e)
        else:
            rest.append(e)
    return (ExceptionGroup(eg.message, matched), ExceptionGroup(eg.message, rest))
"#;

/// Pure-Python `logging` module. `import logging` re-binds `logging` to this
/// object (the import is a recognized no-op for stdlib names). Level constants
/// resolve via [namespace_constants].
const LOGGING_PRELUDE: &str = r#"
class LogRecord:
    def __init__(self, *a, **k):
        pass
class Formatter:
    def __init__(self, *a, **k):
        pass
    def format(self, record):
        return ""
class Filter:
    def __init__(self, *a, **k):
        pass
class Handler:
    def __init__(self, *a, **k):
        self.level = 0
    def setLevel(self, l):
        self.level = l
    def setFormatter(self, f):
        pass
class StreamHandler(Handler):
    pass
class FileHandler(Handler):
    pass
class __PyLogger:
    def __init__(self, name):
        self.name = name
        self.level = 0
        self.handlers = []
    def setLevel(self, l):
        self.level = l
    def debug(self, *a, **k):
        pass
    def info(self, *a, **k):
        pass
    def warning(self, *a, **k):
        pass
    def error(self, *a, **k):
        pass
    def critical(self, *a, **k):
        pass
    def exception(self, *a, **k):
        pass
    def log(self, *a, **k):
        pass
    def isEnabledFor(self, l):
        return True
    def hasHandlers(self):
        return len(self.handlers) > 0
    def addHandler(self, h):
        self.handlers.append(h)
class __LoggingModule:
    def __init__(self):
        self.LogRecord = LogRecord
        self.Formatter = Formatter
        self.Filter = Filter
        self.Handler = Handler
        self.StreamHandler = StreamHandler
        self.FileHandler = FileHandler
        self.NOTSET = 0
        self.DEBUG = 10
        self.INFO = 20
        self.WARNING = 30
        self.WARN = 30
        self.ERROR = 40
        self.CRITICAL = 50
        self.FATAL = 50
        self.lastResort = StreamHandler()
        self.root = __PyLogger("root")
    def getLogger(self, name="root"):
        return __PyLogger(name)
    def basicConfig(self, *a, **k):
        pass
    def getLevelName(self, level):
        names = {0: "NOTSET", 10: "DEBUG", 20: "INFO", 30: "WARNING", 40: "ERROR", 50: "CRITICAL"}
        if level in names:
            return names[level]
        return "Level " + str(level)
    def addLevelName(self, level, name):
        pass
    def debug(self, *a, **k):
        pass
    def info(self, *a, **k):
        pass
    def warning(self, *a, **k):
        pass
    def error(self, *a, **k):
        pass
    def critical(self, *a, **k):
        pass
    def log(self, *a, **k):
        pass
logging = __LoggingModule()
"#;

/// Pure-Python `warnings` module.
const WARNINGS_PRELUDE: &str = r#"
class Warning(Exception):
    pass
class UserWarning(Warning):
    pass
class DeprecationWarning(Warning):
    pass
class PendingDeprecationWarning(Warning):
    pass
class SyntaxWarning(Warning):
    pass
class RuntimeWarning(Warning):
    pass
class FutureWarning(Warning):
    pass
class ImportWarning(Warning):
    pass
class UnicodeWarning(Warning):
    pass
class BytesWarning(Warning):
    pass
class ResourceWarning(Warning):
    pass
class __WarningRecord:
    def __init__(self, message, category):
        self.message = message
        self.category = category
class __CatchWarnings:
    def __init__(self, mod, record):
        self.mod = mod
        self.record = record
        self.entries = []
    def __enter__(self):
        if self.record:
            self.mod.recording = self.entries
            return self.entries
        return None
    def __exit__(self, *a):
        self.mod.recording = None
        return False
class __WarningsModule:
    def __init__(self):
        self.recording = None
    def warn(self, message, category=None, *a, **k):
        if self.recording is not None:
            cat = category
            if cat is None:
                cat = UserWarning
            self.recording.append(__WarningRecord(message, cat))
    def warn_explicit(self, *a, **k):
        pass
    def showwarning(self, *a, **k):
        pass
    def formatwarning(self, *a, **k):
        return ""
    def _filters_mutated(self, *a, **k):
        pass
    def filterwarnings(self, *a, **k):
        pass
    def simplefilter(self, *a, **k):
        pass
    def resetwarnings(self):
        pass
    def catch_warnings(self, record=False):
        return __CatchWarnings(self, record)
warnings = __WarningsModule()"#;

/// Pure-Python `traceback` module. Formatting returns plausible non-empty output;
/// content that needs the live exception (`format_exc` → the exception type name)
/// is a best-effort header since `sys.exc_info()` isn't fully populated.
const TRACEBACK_PRELUDE: &str = r#"
class FrameSummary:
    def __init__(self, *a, **k):
        pass
class StackSummary(list):
    pass
class TracebackException:
    def __init__(self, *a, **k):
        pass
    def format(self, *a, **k):
        return ["Traceback (most recent call last):\n"]
class __TracebackModule:
    def __init__(self):
        self.FrameSummary = FrameSummary
        self.StackSummary = StackSummary
        self.TracebackException = TracebackException
    def format_exc(self, *a, **k):
        return "Traceback (most recent call last):\n"
    def format_exception(self, *a, **k):
        return ["Traceback (most recent call last):\n"]
    def format_exception_only(self, *a, **k):
        return ["\n"]
    def format_tb(self, *a, **k):
        return ["  File \"<unknown>\"\n"]
    def format_stack(self, *a, **k):
        return ["  File \"<unknown>\"\n"]
    def extract_tb(self, *a, **k):
        return StackSummary()
    def extract_stack(self, *a, **k):
        return StackSummary()
    def print_exc(self, *a, **k):
        f = k.get("file", None)
        if f is not None:
            f.write("Traceback (most recent call last):\n")
    def print_tb(self, *a, **k):
        pass
    def print_stack(self, *a, **k):
        pass
    def print_exception(self, *a, **k):
        pass
    def clear_frames(self, *a, **k):
        pass
    def walk_tb(self, *a, **k):
        return iter([])
    def walk_stack(self, *a, **k):
        return iter([])
traceback = __TracebackModule()
"#;

/// `contextlib` helpers. Defined as globals so `from contextlib import X` binds
/// them, and also collected on a `contextlib` module object.
const CONTEXTLIB_PRELUDE: &str = r#"
class __NullContext:
    def __init__(self, enter_result=None):
        self.enter_result = enter_result
    def __enter__(self):
        return self.enter_result
    def __exit__(self, *a):
        return False
def nullcontext(enter_result=None):
    return __NullContext(enter_result)
class __Closing:
    def __init__(self, thing):
        self.thing = thing
    def __enter__(self):
        return self.thing
    def __exit__(self, *a):
        __c = getattr(self.thing, "close")
        __c()
        return False
def closing(thing):
    return __Closing(thing)
class __Suppress:
    def __init__(self, *exc):
        self.exc = exc
    def __enter__(self):
        return None
    def __exit__(self, exc_type, *a):
        if exc_type is None:
            return False
        for e in self.exc:
            if issubclass(exc_type, e):
                return True
        return False
def suppress(*exc):
    return __Suppress(*exc)
class __GenCM:
    def __init__(self, gen):
        self.gen = gen
    def __enter__(self):
        return next(self.gen)
    def __exit__(self, *a):
        try:
            next(self.gen)
        except StopIteration:
            pass
        return False
def contextmanager(func):
    def __cm_helper(*a, **k):
        return __GenCM(func(*a, **k))
    return __cm_helper
class __RedirectStdout:
    def __init__(self, target):
        self.target = target
    def __enter__(self):
        return self.target
    def __exit__(self, *a):
        return False
def redirect_stdout(target):
    return __RedirectStdout(target)
def redirect_stderr(target):
    return __RedirectStdout(target)
class __ContextlibModule:
    def __init__(self):
        self.nullcontext = nullcontext
        self.closing = closing
        self.suppress = suppress
        self.contextmanager = contextmanager
        self.redirect_stdout = redirect_stdout
        self.redirect_stderr = redirect_stderr
contextlib = __ContextlibModule()
"#;

/// `io.StringIO` — an in-memory text buffer implemented in pure Python (string
/// buffer + cursor). No host I/O; `print(file=buf)` writes here via `.write`.
/// `socket` — the stateful half of the module: the `socket` class, plus the
/// module object that carries the constants and the module-level functions.
///
/// A prelude rather than adapters because a socket IS a stateful object: it
/// holds the WASI resource, the input/output streams, the timeout and the
/// option table across calls. The individual WASI calls it sequences are all
/// plain `host:` profile entries (`_wasi_*`), so no bytecode is hand-written
/// here — this only composes them into CPython's blocking API.
///
/// The module object has to carry the module FUNCTIONS too: a global named
/// `socket` shadows the dotted `socket.gethostname` profile builtin, so each
/// one delegates to the bare alias of the same adapter.
const SOCKET_PRELUDE: &str = r#"
class VybeSocketImpl:
    def __init__(self, family=2, kind=1, proto=0, res=None, rx=None, tx=None):
        self.family = family
        self.sock_kind = kind
        self.proto = proto
        self._timeout = None
        self._opts = {}
        self._closed = False
        self._rx = rx
        self._tx = tx
        if res is not None:
            self._res = res
        elif kind == 2:
            self._res = _wasi_udp_new("ipv4")
        else:
            self._res = _wasi_tcp_new("ipv4")
    def settimeout(self, value):
        self._timeout = value
    def gettimeout(self):
        return self._timeout
    def setblocking(self, flag):
        if flag:
            self._timeout = None
        else:
            self._timeout = 0.0
    def setsockopt(self, level, option, value):
        self._opts[str(level) + "/" + str(option)] = value
    def getsockopt(self, level, option, buflen=0):
        key = str(level) + "/" + str(option)
        if key in self._opts:
            return self._opts[key]
        return 0
    def _addr_text(self, address):
        return str(address[0]) + ":" + str(address[1])
    def _addr_tuple(self, record):
        if record is None:
            return ("0.0.0.0", 0)
        parts = record["address"]
        if isinstance(parts, str):
            host = parts
        else:
            pieces = []
            for octet in parts:
                pieces.append(str(octet))
            host = ".".join(pieces)
        return (host, int(record["port"]))
    def bind(self, address):
        _wasi_start_bind(self._res, _wasi_network(), self._addr_text(address))
        _wasi_finish_bind(self._res)
    def listen(self, backlog=5):
        _wasi_backlog(self._res, backlog)
        _wasi_start_listen(self._res)
    def getsockname(self):
        return self._addr_tuple(_wasi_local_addr(self._res))
    def getpeername(self):
        return self._addr_tuple(_wasi_remote_addr(self._res))
    def fileno(self):
        return 0
    def accept(self):
        result = _wasi_accept(self._res)
        if result is None:
            return (None, ("0.0.0.0", 0))
        conn = VybeSocketImpl(self.family, self.sock_kind, 0, result[0], result[1], result[2])
        return (conn, conn.getpeername())
    def connect(self, address):
        _wasi_start_conn(self._res, _wasi_network(), self._addr_text(address))
        streams = _wasi_finish_conn(self._res)
        if streams is not None:
            self._rx = streams[0]
            self._tx = streams[1]
    def send(self, data):
        _wasi_stream_write(self._tx, data)
        return len(data)
    def sendall(self, data):
        _wasi_stream_write(self._tx, data)
        return None
    def recv(self, bufsize=1024):
        return _wasi_stream_read(self._rx, bufsize)
    def shutdown(self, how=2):
        _wasi_sock_shutdown(self._res, how)
    def close(self):
        if not self._closed:
            self._closed = True
            _wasi_sock_shutdown(self._res, 2)
    def dup(self):
        return VybeSocketImpl(self.family, self.sock_kind, 0, self._res, self._rx, self._tx)
    def detach(self):
        self._closed = True
        return 0
    def makefile(self, mode="r", buffering=-1):
        return self
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False
def create_connection(address, timeout=None):
    conn = VybeSocketImpl(2, 1)
    if timeout is not None:
        conn.settimeout(timeout)
    conn.connect(address)
    return conn

class VybeSocketTimeout(OSError):
    pass

class VybeSocketGaiError(OSError):
    pass

def inet_pton(family, text):
    return inet_aton(text)

def inet_ntop(family, packed):
    return inet_ntoa(packed)

def _vybe_swap32(value):
    parts = _vybe_ip4_octets(value)
    return parts[3] * 16777216 + parts[2] * 65536 + parts[1] * 256 + parts[0]

def ntohl(value):
    return _vybe_swap32(value)

def htonl(value):
    return _vybe_swap32(value)

def ntohs(value):
    low = int(value / 256)
    return (value - low * 256) * 256 + low

def htons(value):
    return ntohs(value)

def getdefaulttimeout():
    return None

def setdefaulttimeout(value):
    return None

class VybeSocketModule:
    AF_INET = 2
    AF_INET6 = 10
    AF_UNIX = 1
    SOCK_STREAM = 1
    SOCK_DGRAM = 2
    SOL_SOCKET = 1
    SO_REUSEADDR = 2
    SO_KEEPALIVE = 9
    SO_BROADCAST = 6
    IPPROTO_TCP = 6
    IPPROTO_UDP = 17
    SHUT_RD = 0
    SHUT_WR = 1
    SHUT_RDWR = 2
    has_ipv6 = True
    socket = VybeSocketImpl
    timeout = VybeSocketTimeout
    error = OSError
    gaierror = VybeSocketGaiError
    create_connection = create_connection
    gethostname = gethostname
    gethostbyname = gethostbyname
    getaddrinfo = getaddrinfo
    getservbyname = getservbyname
    inet_aton = inet_aton
    inet_ntoa = inet_ntoa
    inet_pton = inet_pton
    inet_ntop = inet_ntop
    ntohl = ntohl
    htonl = htonl
    ntohs = ntohs
    htons = htons
    getdefaulttimeout = getdefaulttimeout
    setdefaulttimeout = setdefaulttimeout

socket = VybeSocketModule()
"#;

/// `ipaddress` — pure address arithmetic, so no host surface is involved at
/// all; the module is a prelude only because `IPv4Address`/`IPv4Network` are
/// stateful classes.
///
/// Authoring constraints this file must respect (each one silently drops the
/// whole prelude): no comment inside an indented block, no blank line inside a
/// class body, no string literal containing `:` inside a subscript, and no
/// parameter or attribute named `type`.
const IPADDRESS_PRELUDE: &str = r#"
class VybeIPv4Address:
    def __init__(self, value):
        self.version = 4
        self._int = value
        self._text = _vybe_ip4_str(value)
    def __str__(self):
        return self._text
    def __repr__(self):
        return "IPv4Address('" + self._text + "')"
    def __int__(self):
        return self._int
    def __eq__(self, other):
        return int(self) == int(other)
    def __add__(self, n):
        return VybeIPv4Address(self._int + n)
    def __sub__(self, n):
        return VybeIPv4Address(self._int - n)
    def _octets(self):
        return _vybe_ip4_octets(self._int)
    def _prop_compressed(self):
        return self._text
    def _prop_exploded(self):
        return self._text
    def _prop_packed(self):
        return bytes(self._octets())
    def _prop_is_private(self):
        parts = self._octets()
        if parts[0] == 10:
            return True
        if parts[0] == 127:
            return True
        if parts[0] == 192:
            if parts[1] == 168:
                return True
        if parts[0] == 172:
            if parts[1] >= 16:
                if parts[1] <= 31:
                    return True
        return False
    def _prop_is_loopback(self):
        return self._octets()[0] == 127
    def _prop_is_multicast(self):
        first = self._octets()[0]
        if first < 224:
            return False
        return first <= 239
    def _prop_is_global(self):
        return not self._prop_is_private()

class VybeIPv6Address:
    def __init__(self, text):
        self.version = 6
        self._groups = _vybe_ip6_groups(text)
    def __str__(self):
        return _vybe_ip6_compress(self._groups)
    def __repr__(self):
        return "IPv6Address('" + str(self) + "')"
    def __eq__(self, other):
        return str(self) == str(other)
    def _prop_compressed(self):
        return _vybe_ip6_compress(self._groups)
    def _prop_exploded(self):
        return _vybe_ip6_explode(self._groups)
    def _prop_is_loopback(self):
        return _vybe_ip6_explode(self._groups) == "0000:0000:0000:0000:0000:0000:0000:0001"
    def _prop_is_multicast(self):
        return self._groups[0] >= 65280
    def _prop_is_private(self):
        return self._groups[0] >= 64512
    def _prop_is_global(self):
        return not self._prop_is_private()

class VybeIPv4Network:
    def __init__(self, text):
        pair = _vybe_ip4_net_parts(text)
        self.version = 4
        self.prefixlen = pair[1]
        self._mask = _vybe_ip4_mask(pair[1])
        self.num_addresses = _vybe_ip4_count(pair[1])
        self._base = int(pair[0] / self.num_addresses) * self.num_addresses
        self.network_address = VybeIPv4Address(self._base)
        self.netmask = VybeIPv4Address(self._mask)
        self.hostmask = VybeIPv4Address(4294967295 - self._mask)
        self.broadcast_address = VybeIPv4Address(self._base + self.num_addresses - 1)
    def __str__(self):
        return str(self.network_address) + "/" + str(self.prefixlen)
    def __repr__(self):
        return "IPv4Network('" + str(self) + "')"
    def __contains__(self, addr):
        value = int(addr)
        if value < self._base:
            return False
        return value < self._base + self.num_addresses
    def hosts(self):
        out = []
        i = self._base + 1
        last = self._base + self.num_addresses - 1
        while i < last:
            out.append(VybeIPv4Address(i))
            i = i + 1
        return out
    def subnets(self, prefixlen_diff=1):
        new_len = self.prefixlen + prefixlen_diff
        step = _vybe_ip4_count(new_len)
        out = []
        i = self._base
        limit = self._base + self.num_addresses
        while i < limit:
            out.append(VybeIPv4Network(_vybe_ip4_str(i) + "/" + str(new_len)))
            i = i + step
        return out
    def supernet(self, prefixlen_diff=1):
        new_len = self.prefixlen - prefixlen_diff
        return VybeIPv4Network(_vybe_ip4_str(self._base) + "/" + str(new_len))

class VybeIPv4Interface:
    def __init__(self, text):
        pair = _vybe_ip4_net_parts(text)
        self.version = 4
        self.ip = VybeIPv4Address(pair[0])
        self.network = VybeIPv4Network(text)
        self.prefixlen = pair[1]
    def __str__(self):
        return str(self.ip) + "/" + str(self.prefixlen)

def ip_address(value):
    if isinstance(value, int):
        return VybeIPv4Address(value)
    if isinstance(value, bytes):
        total = 0
        for b in value:
            total = total * 256 + b
        return VybeIPv4Address(total)
    text = str(value)
    if "." in text:
        return VybeIPv4Address(_vybe_ip4_parse(text))
    return VybeIPv6Address(text)

def ip_network(value, strict=True):
    return VybeIPv4Network(str(value))

def ip_interface(value):
    return VybeIPv4Interface(str(value))

def collapse_addresses(nets):
    return list(nets)
"#;

/// `http.client` and `ssl` — module SURFACE only.
///
/// Neither maps onto `wasi:sockets` or `ecma:*`, because neither does any
/// networking here: what programs read from these modules is the status-code
/// table, the connection and exception CLASSES, and the protocol constants.
/// The transport, when one is needed, is the `socket` class above.
const HTTP_SSL_PRELUDE: &str = r#"
class VybeHTTPMessage:
    def __init__(self, headers=None):
        self._headers = headers if headers is not None else {}
    def get(self, name, default=None):
        key = str(name).lower()
        if key in self._headers:
            return self._headers[key]
        return default
    def items(self):
        return list(self._headers.items())
    def keys(self):
        return list(self._headers.keys())

class VybeHTTPResponse:
    def __init__(self, status=200, reason="OK", body=""):
        self.status = status
        self.reason = reason
        self._body = body
        self.headers = VybeHTTPMessage()
    def read(self, amt=-1):
        return self._body
    def getheader(self, name, default=None):
        return self.headers.get(name, default)
    def getheaders(self):
        return self.headers.items()
    def close(self):
        return None

class VybeHTTPConnection:
    def __init__(self, host, port=80, timeout=None):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.sock = None
        self._response = None
    def connect(self):
        self.sock = VybeSocketImpl(2, 1)
        self.sock.connect((self.host, self.port))
    def request(self, method, url, body=None, headers=None):
        self._response = VybeHTTPResponse(200, "OK", "")
    def getresponse(self):
        if self._response is None:
            self._response = VybeHTTPResponse(200, "OK", "")
        return self._response
    def close(self):
        if self.sock is not None:
            self.sock.close()

class VybeHTTPSConnection(VybeHTTPConnection):
    def __init__(self, host, port=443, timeout=None):
        VybeHTTPConnection.__init__(self, host, port, timeout)

class VybeHTTPException(Exception):
    pass

class VybeBadStatusLine(VybeHTTPException):
    pass

class VybeIncompleteRead(VybeHTTPException):
    pass

def parse_headers(fp):
    return VybeHTTPMessage()

class VybeSSLContext:
    def __init__(self, protocol=2):
        self.protocol = protocol
        self.verify_mode = 0
        self.check_hostname = False
    def get_ciphers(self):
        return []
    def set_ciphers(self, spec):
        return None
    def load_verify_locations(self, cafile=None, capath=None, cadata=None):
        return None
    def load_default_certs(self, purpose=None):
        return None
    def wrap_socket(self, sock, server_hostname=None, server_side=False):
        return sock

class VybeSSLError(OSError):
    pass

class VybeCertificateError(VybeSSLError):
    pass

class VybeTLSVersion:
    TLSv1 = 769
    TLSv1_1 = 770
    TLSv1_2 = 771
    TLSv1_3 = 772

class VybeSSLPurpose:
    SERVER_AUTH = "serverAuth"
    CLIENT_AUTH = "clientAuth"

def create_default_context(purpose=None, cafile=None, capath=None, cadata=None):
    return VybeSSLContext(2)

def ssl_wrap_socket(sock, keyfile=None, certfile=None, server_side=False):
    return sock

def match_hostname(cert, hostname):
    return None

def enum_certificates(store_name="ROOT"):
    return []

class VybeHttpClientModule:
    OK = 200
    CREATED = 201
    ACCEPTED = 202
    NO_CONTENT = 204
    MOVED_PERMANENTLY = 301
    FOUND = 302
    NOT_MODIFIED = 304
    BAD_REQUEST = 400
    UNAUTHORIZED = 401
    FORBIDDEN = 403
    NOT_FOUND = 404
    METHOD_NOT_ALLOWED = 405
    REQUEST_TIMEOUT = 408
    CONFLICT = 409
    GONE = 410
    INTERNAL_SERVER_ERROR = 500
    NOT_IMPLEMENTED = 501
    BAD_GATEWAY = 502
    SERVICE_UNAVAILABLE = 503
    GATEWAY_TIMEOUT = 504
    HTTPConnection = VybeHTTPConnection
    HTTPSConnection = VybeHTTPSConnection
    HTTPResponse = VybeHTTPResponse
    HTTPMessage = VybeHTTPMessage
    HTTPException = VybeHTTPException
    BadStatusLine = VybeBadStatusLine
    IncompleteRead = VybeIncompleteRead
    parse_headers = parse_headers
    responses = {200: "OK", 201: "Created", 202: "Accepted", 204: "No Content", 301: "Moved Permanently", 302: "Found", 304: "Not Modified", 400: "Bad Request", 401: "Unauthorized", 403: "Forbidden", 404: "Not Found", 405: "Method Not Allowed", 408: "Request Timeout", 409: "Conflict", 410: "Gone", 500: "Internal Server Error", 501: "Not Implemented", 502: "Bad Gateway", 503: "Service Unavailable", 504: "Gateway Timeout"}

class VybeHttpModule:
    client = VybeHttpClientModule()

class VybeSslModule:
    CERT_NONE = 0
    CERT_OPTIONAL = 1
    CERT_REQUIRED = 2
    PROTOCOL_TLS = 2
    PROTOCOL_TLS_CLIENT = 16
    PROTOCOL_TLS_SERVER = 17
    OP_NO_SSLv2 = 16777216
    OP_NO_SSLv3 = 33554432
    HAS_TLSv1_3 = True
    SSLContext = VybeSSLContext
    SSLError = VybeSSLError
    CertificateError = VybeCertificateError
    TLSVersion = VybeTLSVersion
    Purpose = VybeSSLPurpose
    create_default_context = create_default_context
    wrap_socket = ssl_wrap_socket
    match_hostname = match_hostname
    enum_certificates = enum_certificates

http = VybeHttpModule()
ssl = VybeSslModule()
"#;

const IO_PRELUDE: &str = r#"
class StringIO:
    def __init__(self, initial=''):
        self._parts = []
        self._pos = 0
        self.closed = False
        if isinstance(initial, str) and initial != '':
            self._parts.append(initial)
    def write(self, s):
        self._parts.append(s)
        self._pos = self._pos + len(s)
        return len(s)
    def writelines(self, lines):
        for line in lines:
            self.write(line)
    def read(self, size=-1):
        data = ''.join(self._parts)
        pos = self._pos
        if size is None or size < 0:
            result = data[pos:]
        else:
            result = data[pos:pos + size]
        self._pos = pos + len(result)
        return result
    def readline(self):
        data = ''.join(self._parts)
        pos = self._pos
        rest = data[pos:]
        if rest == '':
            return ''
        idx = rest.find(chr(10))
        if idx < 0:
            result = rest
        else:
            result = rest[:idx + 1]
        self._pos = pos + len(result)
        return result
    def readlines(self):
        data = ''.join(self._parts)
        pos = self._pos
        rest = data[pos:]
        self._pos = len(data)
        result = []
        if rest == '':
            return result
        parts = rest.split(chr(10))
        n = len(parts)
        i = 0
        while i < n:
            p = parts[i]
            if i < n - 1:
                result.append(p + chr(10))
            elif p != '':
                result.append(p)
            i += 1
        return result
    def getvalue(self):
        return ''.join(self._parts)
    def seek(self, pos, whence=0):
        if whence == 1:
            self._pos = self._pos + pos
        elif whence == 2:
            self._pos = len(''.join(self._parts)) + pos
        else:
            self._pos = pos
        return self._pos
    def tell(self):
        return self._pos
    def truncate(self, size=None):
        end = size
        if end is None:
            end = self._pos
        data = ''.join(self._parts)
        self._parts = [data[:end]]
        return end
    def __iter__(self):
        return iter(self.readlines())
    def readable(self):
        return True
    def writable(self):
        return True
    def seekable(self):
        return True
    def flush(self):
        pass
    def close(self):
        self.closed = True
    def detach(self):
        return None
    def __enter__(self):
        return self
    def __exit__(self, *a):
        self.close()
        return False
class BytesIO:
    def __init__(self, initial=b''):
        self._parts = []
        self._pos = 0
        self.closed = False
        if isinstance(initial, bytes) and len(initial) != 0:
            self._parts.append(initial)
            self._pos = len(initial)
    def write(self, b):
        self._parts.append(b)
        self._pos += len(b)
        return len(b)
    def read(self, size=-1):
        data = b''.join(self._parts)
        pos = self._pos
        if size is None or size < 0:
            result = data[pos:]
        else:
            result = data[pos:pos + size]
        self._pos = pos + len(result)
        return result
    def read1(self, size=-1):
        return self.read(size)
    def getvalue(self):
        return b''.join(self._parts)
    def getbuffer(self):
        return b''.join(self._parts)
    def seek(self, pos, whence=0):
        if whence == 1:
            self._pos = self._pos + pos
        elif whence == 2:
            self._pos = len(b''.join(self._parts)) + pos
        else:
            self._pos = pos
        return self._pos
    def tell(self):
        return self._pos
    def readable(self):
        return True
    def writable(self):
        return True
    def seekable(self):
        return True
    def flush(self):
        pass
    def close(self):
        self.closed = True
    def __iter__(self):
        return iter(b''.join(self._parts))
    def __enter__(self):
        return self
    def __exit__(self, *a):
        self.close()
        return False
class __IOModule:
    def __init__(self):
        self.StringIO = StringIO
        self.BytesIO = BytesIO
io = __IOModule()
"#;

/// `fnmatch` — shell-style pattern matching.
///
/// The MATCHER is no longer here: `__glob_match` is bound to
/// `common:str_glob_match`, the shared emitter php's `fnmatch` also uses. What
/// was here was a hand-written iterative `*`/`?` matcher that could not do
/// `[seq]` classes at all — so `fnmatch("fileA.py", "file[ABC].py")` was False
/// where real python3 says True. Every test missed it because
/// `fold_fnmatch_call` CONSTANT-FOLDS literal arguments in Rust, so the prelude
/// only ran behind a variable, which no test exercised.
///
/// `os.path.normcase` is identity on POSIX, so `fnmatch` is case-SENSITIVE here
/// and identical to `fnmatchcase` — measured, `fnmatch.fnmatch('ABC','abc')` is
/// False on this platform. The old body lower-cased both sides.
///
/// `translate` stays Python: it emits python's OWN regex dialect
/// (`(?s:…)\Z`), which is a regex-layer question, not a glob one.
const FNMATCH_PRELUDE: &str = r#"
def __fn_translate(pat):
    res = ''
    i = 0
    n = len(pat)
    while i < n:
        c = pat[i]
        if c == '*':
            res = res + '.*'
        elif c == '?':
            res = res + '.'
        elif c == '.' or c == '\\' or c == '+' or c == '(' or c == ')' or c == '|' or c == '^' or c == '$' or c == '{' or c == '}':
            res = res + '\\' + c
        else:
            res = res + c
        i = i + 1
    return '(?s:' + res + ')\\Z'
class __FnmatchModule:
    def fnmatch(self, name, pat):
        return __glob_match(name, pat)
    def fnmatchcase(self, name, pat):
        return __glob_match(name, pat)
    def filter(self, names, pat):
        result = []
        for nm in names:
            if __glob_match(nm, pat):
                result.append(nm)
        return result
    def translate(self, pat):
        return __fn_translate(pat)
fnmatch = __FnmatchModule()
"#;

/// `configparser` — a minimal INI parser. Data lives in nested dicts (dict
/// attrs behave; string manipulation is on locals, dodging the self-attr string
/// slice/concat bug).
const CONFIGPARSER_PRELUDE: &str = r#"
class ConfigParser:
    def __init__(self, defaults=None, dict_type=None, allow_no_value=False):
        self._sections = {}
    def read_string(self, s, source='<string>'):
        current = None
        hash_c = chr(35)
        semi_c = chr(59)
        lbrack = chr(91)
        lines = s.split(chr(10))
        for line in lines:
            t = line.strip()
            blank = t == ''
            comment = not blank and (t[0] == hash_c or t[0] == semi_c)
            header = not blank and not comment and t[0] == lbrack
            if header:
                name = t[1:len(t) - 1]
                secs = self._sections
                secs[name] = {}
                self._sections = secs
                current = name
            elif not blank and not comment and current is not None:
                idx = t.find('=')
                if idx < 0:
                    idx = t.find(':')
                if idx >= 0:
                    key = t[:idx].strip()
                    val = t[idx + 1:].strip()
                    secs = self._sections
                    inner = secs[current]
                    inner[key] = val
                    secs[current] = inner
                    self._sections = secs
    def read_dict(self, d):
        for name in d:
            target = {}
            src = d[name]
            for key in src:
                target[key] = str(src[key])
            self._sections[name] = target
    def sections(self):
        result = []
        for k in self._sections:
            result.append(k)
        return result
    def has_section(self, sec):
        return sec in self._sections
    def has_option(self, sec, opt):
        return sec in self._sections and opt in self._sections[sec]
    def options(self, sec):
        result = []
        for k in self._sections[sec]:
            result.append(k)
        return result
    def get(self, sec, opt, fallback=None):
        if sec in self._sections and opt in self._sections[sec]:
            return self._sections[sec][opt]
        return fallback
    def getint(self, sec, opt):
        return int(self._sections[sec][opt])
    def getfloat(self, sec, opt):
        return float(self._sections[sec][opt])
    def getboolean(self, sec, opt):
        v = self._sections[sec][opt].lower()
        return v == 'true' or v == '1' or v == 'yes' or v == 'on'
    def defaults(self):
        return {}
    def __getitem__(self, sec):
        return self._sections[sec]
    def __contains__(self, sec):
        return sec in self._sections
class __ConfigparserModule:
    def __init__(self):
        self.ConfigParser = ConfigParser
        self.RawConfigParser = ConfigParser
configparser = __ConfigparserModule()
"#;

/// `pathlib` — PurePath/Path (pure Python). All path math runs on LOCAL
/// strings (dodging the self-attr slice/concat bug); only the finished path
/// string lands in `self._s`. Derived fields are `@property` getters (no
/// stored slices, no parent-construction recursion). String literals contain
/// no `#`/`[` (both break the prelude preprocessor) — `/`, `:`, `.` only.
/// FS predicates (`exists`/`is_dir`/`is_file`) return a conservative `False`:
/// no `wasi:filesystem` binding is resolvable at load, so existence cannot be
/// confirmed without a host FS (out of this layer's reach).
const PATHLIB_PRELUDE: &str = r#"
import os
def _pp_str(p):
    if hasattr(p, '_s'):
        return p._s
    return p
def _pp_norm(s):
    s2 = s.replace(chr(92), '/')
    while s2.find('//') >= 0:
        s2 = s2.replace('//', '/')
    n = len(s2)
    if n > 1 and s2[n - 1] == '/':
        s2 = s2[0:n - 1]
    return s2
def _pp_join_one(base, seg):
    s = _pp_str(seg).replace(chr(92), '/')
    if len(s) > 0 and s[0] == '/':
        return _pp_norm(s)
    if base == '':
        return _pp_norm(s)
    n = len(base)
    if base[n - 1] == '/':
        return _pp_norm(base + s)
    return _pp_norm(base + '/' + s)
def _pp_fnmatch(name, pat):
    ni = 0
    pi = 0
    nlen = len(name)
    plen = len(pat)
    star_pi = -1
    star_ni = 0
    while ni < nlen:
        if pi < plen and (pat[pi] == name[ni] or pat[pi] == '?'):
            ni = ni + 1
            pi = pi + 1
        elif pi < plen and pat[pi] == '*':
            star_pi = pi
            star_ni = ni
            pi = pi + 1
        elif star_pi >= 0:
            pi = star_pi + 1
            star_ni = star_ni + 1
            ni = star_ni
        else:
            return False
    while pi < plen and pat[pi] == '*':
        pi = pi + 1
    return pi == plen
class __PathStat:
    def __init__(self):
        self.st_size = 0
        self.st_mode = 0
        self.st_mtime = 0
        self.st_ctime = 0
        self.st_atime = 0
class PurePath:
    def __init__(self, p=''):
        self._s = _pp_norm(_pp_str(p))
    def _is_win(self):
        return False
    def _make(self, s):
        return PurePath(s)
    @property
    def drive(self):
        if not self._is_win():
            return ''
        s = self._s
        if len(s) >= 2 and s[1] == ':':
            return s[0:2]
        return ''
    @property
    def root(self):
        d = self.drive
        rest = self._s[len(d):]
        if len(rest) >= 1 and rest[0] == '/':
            return '/'
        return ''
    @property
    def anchor(self):
        d = self.drive
        r = self.root
        return d + r
    @property
    def name(self):
        a = self.anchor
        rest = self._s[len(a):]
        last = ''
        for c in rest.split('/'):
            if c != '':
                last = c
        return last
    @property
    def stem(self):
        nm = self.name
        idx = nm.rfind('.')
        if idx > 0:
            return nm[0:idx]
        return nm
    @property
    def suffix(self):
        nm = self.name
        idx = nm.rfind('.')
        if idx > 0:
            return nm[idx:]
        return ''
    @property
    def suffixes(self):
        nm = self.name
        result = []
        if nm.endswith('.'):
            return result
        pieces = nm.split('.')
        i = 1
        while i < len(pieces):
            result.append('.' + pieces[i])
            i = i + 1
        return result
    @property
    def parts(self):
        anchor = self.anchor
        rest = self._s[len(anchor):]
        result = []
        if anchor != '':
            result.append(anchor)
        for c in rest.split('/'):
            if c != '':
                result.append(c)
        return result
    @property
    def parent(self):
        anchor = self.anchor
        rest = self._s[len(anchor):]
        kept = []
        for c in rest.split('/'):
            if c != '':
                kept.append(c)
        if len(kept) <= 1:
            if anchor != '':
                return self._make(anchor)
            return self._make('.')
        newrest = ''
        i = 0
        while i < len(kept) - 1:
            if i > 0:
                newrest = newrest + '/'
            newrest = newrest + kept[i]
            i = i + 1
        return self._make(anchor + newrest)
    def with_name(self, newname):
        prt = self.parent
        ps = prt._s
        if ps == '.' or ps == '':
            return self._make(newname)
        if ps[len(ps) - 1] == '/':
            return self._make(ps + newname)
        return self._make(ps + '/' + newname)
    def with_suffix(self, suf):
        nm = self.name
        idx = nm.rfind('.')
        if idx > 0:
            base = nm[0:idx]
        else:
            base = nm
        return self.with_name(base + suf)
    def with_stem(self, newstem):
        sfx = self.suffix
        return self.with_name(newstem + sfx)
    def match(self, pat):
        nm = self.name
        return _pp_fnmatch(nm, pat)
    def joinpath(self, *others):
        cur = self._s
        for o in others:
            cur = _pp_join_one(cur, o)
        return self._make(cur)
    def __truediv__(self, other):
        return self._make(_pp_join_one(self._s, other))
    def relative_to(self, other):
        o = _pp_norm(_pp_str(other))
        s = self._s
        if s == o:
            return self._make('.')
        prefix = o
        if len(prefix) == 0 or prefix[len(prefix) - 1] != '/':
            prefix = prefix + '/'
        if s[0:len(prefix)] == prefix:
            return self._make(s[len(prefix):])
        return self._make(s)
    def is_relative_to(self, other):
        o = _pp_norm(_pp_str(other))
        s = self._s
        if s == o:
            return True
        prefix = o
        if len(prefix) == 0 or prefix[len(prefix) - 1] != '/':
            prefix = prefix + '/'
        return s[0:len(prefix)] == prefix
    def as_posix(self):
        return self._s
    def as_uri(self):
        s = self._s
        if len(s) == 0 or s[0] != '/':
            p = '/' + s
        else:
            p = s
        return 'file://' + p
    def is_absolute(self):
        r = self.root
        if self._is_win():
            d = self.drive
            return d != '' and r != ''
        return r == '/'
    def is_reserved(self):
        if not self._is_win():
            return False
        nm0 = self.name
        nm = nm0.upper()
        dot = nm.find('.')
        if dot >= 0:
            nm = nm[0:dot]
        reserved = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'LPT1', 'LPT2']
        for r in reserved:
            if nm == r:
                return True
        return False
    @staticmethod
    def from_uri(uri):
        u = uri
        pre = 'file://'
        if u[0:len(pre)] == pre:
            u = u[len(pre):]
        return PurePath(u)
    def __eq__(self, other):
        if not hasattr(other, '_s'):
            return False
        a = self._s
        b = other._s
        return a == b
    def __hash__(self):
        h = 0
        s = self._s
        for ch in s:
            h = h * 31 + ord(ch)
        return h
    def __str__(self):
        return self._s
    def __repr__(self):
        return self._s
class PurePosixPath(PurePath):
    def _is_win(self):
        return False
class PureWindowsPath(PurePath):
    def _is_win(self):
        return True
class Path(PurePath):
    def _make(self, s):
        return Path(s)
    def exists(self):
        return os.path.exists(self._s)
    def is_dir(self):
        return os.path.isdir(self._s)
    def is_file(self):
        return os.path.isfile(self._s)
    def resolve(self, strict=False):
        return self
    def absolute(self):
        return self
    def expanduser(self):
        return self
    @staticmethod
    def home():
        return Path('/home')
    @staticmethod
    def cwd():
        return Path('.')
    def stat(self):
        return __PathStat()
    def lstat(self):
        return __PathStat()
    def iterdir(self):
        return []
    def glob(self, pat):
        return []
    def rglob(self, pat):
        return []
    def read_text(self, encoding=None):
        return ''
    def write_text(self, data, encoding=None):
        return len(data)
    def read_bytes(self):
        return None
    def write_bytes(self, data):
        return 0
    def open(self, mode='r', encoding=None):
        return None
    def touch(self, mode=438, exist_ok=True):
        return None
    def mkdir(self, mode=511, parents=False, exist_ok=False):
        return None
    def rmdir(self):
        return None
    def unlink(self, missing_ok=False):
        return None
    def rename(self, target):
        return Path(_pp_str(target))
    def replace(self, target):
        return Path(_pp_str(target))
    def hardlink_to(self, target):
        return None
    def symlink_to(self, target, target_is_directory=False):
        return None
    def chmod(self, mode):
        return None
    def samefile(self, other):
        return self._s == _pp_str(other)
"#;

/// `random` distributions and range/weight helpers as pure Python over the
/// working `random.random()` entropy plus `math`. Injected when the source
/// references `random`; the walker rewrites the non-host-backed names
/// (`gauss`, `uniform`, `randrange`, `choices`, …) → `__py_random_*`.
const RANDOM_PRELUDE: &str = r#"
import random
import math
def __py_random_r():
    return random.random()
def __py_random_uniform(a, b):
    return a + (b - a) * __py_random_r()
def __py_random_expovariate(lambd):
    return -math.log(1.0 - __py_random_r()) / lambd
def __py_random_gauss(mu, sigma):
    x2pi = __py_random_r() * 2.0 * math.pi
    g2rad = math.sqrt(-2.0 * math.log(1.0 - __py_random_r()))
    return mu + math.cos(x2pi) * g2rad * sigma
def __py_random_normalvariate(mu, sigma):
    return __py_random_gauss(mu, sigma)
def __py_random_lognormvariate(mu, sigma):
    return math.exp(__py_random_normalvariate(mu, sigma))
def __py_random_triangular(low, high):
    u = __py_random_r()
    return low + (high - low) * math.sqrt(u)
def __py_random_paretovariate(alpha):
    u = 1.0 - __py_random_r()
    return 1.0 / math.exp(math.log(u) / alpha)
def __py_random_weibullvariate(alpha, beta):
    u = 1.0 - __py_random_r()
    return alpha * math.exp(math.log(-math.log(u)) / beta)
def __py_random_vonmisesvariate(mu, kappa):
    return mu + (__py_random_r() - 0.5) * kappa
def __py_random_gammavariate(alpha, beta):
    gv_total = 0.0
    gv_i = 0
    while gv_i < 3:
        gv_total = gv_total + (-math.log(1.0 - __py_random_r()))
        gv_i = gv_i + 1
    return gv_total * beta
def __py_random_betavariate(alpha, beta):
    y1 = __py_random_gammavariate(alpha, 1.0)
    y2 = __py_random_gammavariate(beta, 1.0)
    return y1 / (y1 + y2)
def __py_random_getrandbits(k):
    gb_total = 0
    gb_i = 0
    while gb_i < k:
        gb_bit = 0
        if __py_random_r() < 0.5:
            gb_bit = 1
        gb_total = gb_total * 2 + gb_bit
        gb_i = gb_i + 1
    return gb_total
def __py_random_randbytes(n):
    rb_vals = []
    rb_i = 0
    while rb_i < n:
        rb_vals.append(__py_random_getrandbits(8))
        rb_i = rb_i + 1
    return bytes(rb_vals)
def __py_random_randrange(start, stop, step):
    if step == 0:
        raise ValueError('randrange step must not be zero')
    if step > 0:
        n = (stop - start + step - 1) // step
    else:
        n = (start - stop - step - 1) // (0 - step)
    if n <= 0:
        raise ValueError('empty range for randrange')
    idx = int(__py_random_r() * n)
    if idx >= n:
        idx = n - 1
    return start + idx * step
def __py_random_choices(pop, weights, cum, k):
    result = []
    total = 0.0
    cums = []
    if cum is not None:
        for w in cum:
            cums.append(w * 1.0)
        total = cums[len(cums) - 1]
    elif weights is not None:
        acc = 0.0
        for w in weights:
            acc = acc + w
            cums.append(acc)
        total = acc
    else:
        j = 0
        while j < len(pop):
            cums.append(j + 1.0)
            j = j + 1
        total = len(pop) * 1.0
    c = 0
    while c < k:
        target = __py_random_r() * total
        pick = len(pop) - 1
        m = 0
        chosen = False
        while m < len(cums):
            if not chosen and target < cums[m]:
                pick = m
                chosen = True
            m = m + 1
        result.append(pop[pick])
        c = c + 1
    return result
def __py_random_getstate():
    return (0, 0, 0)
def __py_random_setstate(state):
    return None
"#;

/// Pure-Python `string` module surface: `Template` (`$name`/`${name}`
/// substitution with `$$` escape and a class-attribute `delimiter`),
/// `Formatter` (delegates to `str.format`), and `capwords`. Constants
/// (ascii_letters, digits, …) are intercepted in [desugar_member_reads].
/// Pure-Python `pprint` surface: `pformat`/`pprint`/`pp`, the stateful
/// `PrettyPrinter` class, plus `saferepr`/`isreadable`/`isrecursive`.
///
/// A prelude (not adapters) because `PrettyPrinter` is a genuinely stateful
/// CLASS carrying indent/width/depth/compact/sort_dicts across calls, and the
/// module functions are thin wrappers over the same recursive layout routine —
/// splitting them would duplicate the algorithm. See
/// `feedback_adapters_over_preludes`: "a class needs a class".
///
/// Recursion detection tracks container identity down the current path only, so
/// a value repeated in sibling positions is not a cycle.
const PPRINT_PRELUDE: &str = r#"
def __pprint_is_container(o):
    return isinstance(o, (list, tuple, dict, set, frozenset))
def __pprint_cycle(o, seen):
    for s in seen:
        if s is o:
            return True
    return False
def __pprint_has_cycle(o, seen):
    if not __pprint_is_container(o):
        return False
    if __pprint_cycle(o, seen):
        return True
    seen = seen + [o]
    if isinstance(o, dict):
        for k in o:
            if __pprint_has_cycle(o[k], seen):
                return True
        return False
    for it in o:
        if __pprint_has_cycle(it, seen):
            return True
    return False
def __pprint_kind(o):
    if isinstance(o, dict):
        return "dict"
    if isinstance(o, tuple):
        return "tuple"
    if isinstance(o, frozenset):
        return "frozenset"
    if isinstance(o, set):
        return "set"
    return "list"
def __pprint_underscore(n):
    s = str(n)
    neg = s.startswith("-")
    if neg:
        s = s[1:]
    out = ""
    c = 0
    i = len(s) - 1
    while i >= 0:
        out = s[i] + out
        c += 1
        if c % 3 == 0 and i > 0:
            out = "_" + out
        i -= 1
    if neg:
        out = "-" + out
    return out
def __pprint_fmt(o, ind, width, depth, compact, sort_dicts, under, level, seen, col):
    if __pprint_is_container(o):
        if __pprint_cycle(o, seen):
            return "<Recursion on " + __pprint_kind(o) + " with id=0>"
        if depth is not None and level >= depth:
            return "..."
    if isinstance(o, bool) or o is None:
        return repr(o)
    if under and isinstance(o, int):
        return __pprint_underscore(o)
    if not __pprint_is_container(o):
        return repr(o)
    seen = seen + [o]
    pad = " " * (col + ind)
    if isinstance(o, dict):
        keys = list(o)
        if sort_dicts:
            keys = sorted(keys)
        parts = []
        for k in keys:
            parts.append(repr(k) + ": " + __pprint_fmt(o[k], ind, width, depth, compact,
                                                       sort_dicts, under, level + 1, seen, col + ind))
        return __pprint_wrap(parts, "{", "}", width, col, pad, False)
    if isinstance(o, (set, frozenset)):
        parts = []
        for it in sorted(o):
            parts.append(__pprint_fmt(it, ind, width, depth, compact, sort_dicts,
                                      under, level + 1, seen, col + ind))
        body = __pprint_wrap(parts, "{", "}", width, col, pad, False)
        if isinstance(o, frozenset):
            if len(parts) == 0:
                return "frozenset()"
            return "frozenset(" + body + ")"
        if len(parts) == 0:
            return "set()"
        return body
    parts = []
    for it in o:
        parts.append(__pprint_fmt(it, ind, width, depth, compact, sort_dicts,
                                  under, level + 1, seen, col + ind))
    if isinstance(o, tuple):
        return __pprint_wrap(parts, "(", ")", width, col, pad, len(parts) == 1)
    return __pprint_wrap(parts, "[", "]", width, col, pad, False)
def __pprint_wrap(parts, open_c, close_c, width, col, pad, trail_comma):
    flat = open_c + ", ".join(parts) + ("," if trail_comma else "") + close_c
    if col + len(flat) <= width or len(parts) <= 1:
        return flat
    return open_c + (",\n" + pad).join(parts) + close_c
def __pprint_pformat(o, indent=1, width=80, depth=None, compact=False, sort_dicts=True,
                     underscore_numbers=False):
    return __pprint_fmt(o, indent, width, depth, compact, sort_dicts,
                        underscore_numbers, 0, [], 0)
def __pprint_pprint(o, stream=None, indent=1, width=80, depth=None, compact=False,
                    sort_dicts=True, underscore_numbers=False):
    text = __pprint_pformat(o, indent, width, depth, compact, sort_dicts, underscore_numbers)
    if stream is None:
        print(text)
    else:
        stream.write(text + "\n")
def __pprint_pp(o, stream=None, **kwargs):
    __pprint_pprint(o, stream, **kwargs)
def __pprint_saferepr(o):
    return __pprint_fmt(o, 1, 1000000, None, False, True, False, 0, [], 0)
def __pprint_isrecursive(o):
    return __pprint_has_cycle(o, [])
def __pprint_isreadable(o):
    return not __pprint_has_cycle(o, [])
class __pprint_PrettyPrinter:
    def __init__(self, indent=1, width=80, depth=None, stream=None, compact=False,
                 sort_dicts=True, underscore_numbers=False):
        self.indent = indent
        self.width = width
        self.depth = depth
        self.stream = stream
        self.compact = compact
        self.sort_dicts = sort_dicts
        self.underscore_numbers = underscore_numbers
    def pformat(self, o):
        return __pprint_pformat(o, self.indent, self.width, self.depth, self.compact,
                                self.sort_dicts, self.underscore_numbers)
    def pprint(self, o):
        __pprint_pprint(o, self.stream, self.indent, self.width, self.depth,
                        self.compact, self.sort_dicts, self.underscore_numbers)
    def isrecursive(self, o):
        return __pprint_isrecursive(o)
    def isreadable(self, o):
        return __pprint_isreadable(o)
"#;

const STRING_PRELUDE: &str = r#"
def __string_is_id_start(c):
    return c == "_" or ("a" <= c <= "z") or ("A" <= c <= "Z")
def __string_is_id_char(c):
    return c == "_" or ("a" <= c <= "z") or ("A" <= c <= "Z") or ("0" <= c <= "9")
class __string_Template:
    delimiter = "$"
    def __init__(self, template):
        self.template = template
    def _tscan(self, mapping, safe, collect):
        d = self.delimiter
        s = self.template
        out = ""
        ids = []
        i = 0
        n = len(s)
        while i < n:
            c = s[i]
            if c != d:
                out += c
                i += 1
                continue
            if i + 1 < n and s[i + 1] == d:
                out += d
                i += 2
                continue
            if i + 1 < n and s[i + 1] == "{":
                j = i + 2
                name = ""
                while j < n and s[j] != "}":
                    name += s[j]
                    j += 1
                if j < n:
                    if collect:
                        if name not in ids:
                            ids.append(name)
                    elif name in mapping:
                        out += str(mapping[name])
                    elif safe:
                        out += d + "{" + name + "}"
                    else:
                        raise KeyError(name)
                    i = j + 1
                    continue
            if i + 1 < n and __string_is_id_start(s[i + 1]):
                j = i + 1
                name = ""
                while j < n and __string_is_id_char(s[j]):
                    name += s[j]
                    j += 1
                if collect:
                    if name not in ids:
                        ids.append(name)
                elif name in mapping:
                    out += str(mapping[name])
                elif safe:
                    out += d + name
                else:
                    raise KeyError(name)
                i = j
                continue
            out += c
            i += 1
        if collect:
            return ids
        return out
    def _tmerge(self, mapping, kws):
        m = {}
        if mapping is not None:
            for k in mapping:
                m[k] = mapping[k]
        for k in kws:
            m[k] = kws[k]
        return m
    def substitute(self, mapping=None, **kws):
        return self._tscan(self._tmerge(mapping, kws), False, False)
    def safe_substitute(self, mapping=None, **kws):
        return self._tscan(self._tmerge(mapping, kws), True, False)
    def get_identifiers(self):
        return self._tscan({}, False, True)
    def is_valid(self):
        return True
class __string_Formatter:
    def format(self, fmt, *args, **kwargs):
        return fmt.format(*args, **kwargs)
    def vformat(self, fmt, args, kwargs):
        return fmt.format(*args, **kwargs)
def __string_capwords(s, sep=None):
    if sep is None:
        words = s.split()
        joiner = " "
    else:
        words = s.split(sep)
        joiner = sep
    out = []
    for w in words:
        out.append(w.capitalize())
    return joiner.join(out)
"#;

const SHLEX_PRELUDE: &str = r##"
def __py_shlex_needs_quote(s):
    if s == "":
        return True
    safe = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_@%+=:,./-"
    for ch in s:
        if ch not in safe:
            return True
    return False

def __py_shlex_quote(s):
    if not __py_shlex_needs_quote(s):
        return s
    out = "'"
    for ch in s:
        if ch == "'":
            out += "'\"'\"'"
        else:
            out += ch
    return out + "'"

def __py_shlex_split(s, comments=False, posix=True):
    out = []
    cur = ""
    quote = ""
    esc = False
    for ch in s:
        if esc:
            cur += ch
            esc = False
        elif quote != "":
            if ch == quote:
                if posix:
                    quote = ""
                else:
                    cur += ch
                    quote = ""
            else:
                cur += ch
        elif ch == "\\" and posix:
            esc = True
        elif ch == "'" or ch == '"':
            if posix:
                quote = ch
            else:
                quote = ch
                cur += ch
        elif comments and ch == "#":
            break
        elif ch == " " or ch == "\t" or ch == "\n":
            if cur != "":
                out.append(cur)
                cur = ""
        else:
            cur += ch
    if quote != "":
        raise ValueError("No closing quotation")
    if cur != "":
        out.append(cur)
    return out

def __py_shlex_join(parts):
    joined = ""
    first = True
    for p in parts:
        q = __py_shlex_quote(p)
        if first:
            joined = q
            first = False
        else:
            joined += " " + q
    return joined

class __py_shlex_class:
    def __init__(self, instream=None, posix=False, punctuation_chars=False):
        if instream is None:
            self.text = ""
        elif isinstance(instream, str):
            self.text = str(instream)
        else:
            self.text = instream.getvalue()
        self.posix = posix
        self.whitespace_split = False
        self.commenters = "#"
    def __iter__(self):
        text = self.text
        if self.commenters != "":
            lines = []
            for line in text.splitlines():
                lines.append(line.split(self.commenters)[0])
            text = "\n".join(lines)
        return iter(__py_shlex_split(text, posix=self.posix))

def __py_shlex_tokens(obj):
    text = obj.text
    if obj.commenters != "":
        lines = []
        for line in text.splitlines():
            lines.append(line.split(obj.commenters)[0])
        joined = ""
        first = True
        for line in lines:
            if first:
                joined = line
                first = False
            else:
                joined += "\n" + line
        text = joined
    return __py_shlex_split(text, posix=obj.posix)

shlex = {
    "__name__": "shlex",
    "split": __py_shlex_split,
    "quote": __py_shlex_quote,
    "join": __py_shlex_join,
    "shlex": __py_shlex_class }
"##;

const TEXTWRAP_PRELUDE: &str = r##"
def __py_textwrap_words(text):
    return text.expandtabs().split()

def __py_textwrap_wrap(text, width=70, initial_indent="", subsequent_indent="", break_long_words=True, break_on_hyphens=True, expand_tabs=True, replace_whitespace=True, drop_whitespace=True, max_lines=None, placeholder=" [...]"):
    if expand_tabs:
        text = text.expandtabs()
    words = text.split()
    if len(words) == 0:
        return []
    lines = []
    cur = initial_indent
    limit = width
    for word in words:
        if break_long_words and len(word) > width:
            if cur.strip() != "":
                lines.append(cur.rstrip())
                cur = subsequent_indent
            i = 0
            while i < len(word):
                lines.append(word[i:i+width])
                i += width
            continue
        sep = "" if cur == initial_indent or cur == subsequent_indent else " "
        if len(cur) + len(sep) + len(word) <= limit:
            cur += sep + word
        else:
            if cur.strip() != "":
                lines.append(cur.rstrip())
            cur = subsequent_indent + word
    if cur.strip() != "":
        lines.append(cur.rstrip())
    if max_lines is not None and len(lines) > max_lines:
        lines = lines[:max_lines]
        if len(lines) > 0:
            lines[-1] = lines[-1][:max(0, width - len(placeholder))].rstrip() + placeholder
    return lines

def __py_textwrap_fill(text, width=70, initial_indent="", subsequent_indent="", break_long_words=True, break_on_hyphens=True, expand_tabs=True, replace_whitespace=True, drop_whitespace=True, max_lines=None, placeholder=" [...]"):
    return "\n".join(__py_textwrap_wrap(text, width, initial_indent, subsequent_indent, break_long_words, break_on_hyphens, expand_tabs, replace_whitespace, drop_whitespace, max_lines, placeholder))

def __py_textwrap_dedent(text):
    lines = text.split("\n")
    ind = -1
    for line in lines:
        if line.strip() == "":
            continue
        n = len(line) - len(line.lstrip())
        if ind < 0 or n < ind:
            ind = n
    if ind <= 0:
        return text
    out = []
    for line in lines:
        out.append(line[ind:])
    return "\n".join(out)

def __py_textwrap_indent(text, prefix, predicate=None):
    if text == "":
        return ""
    out = []
    for line in text.split("\n"):
        use = True
        if predicate is not None:
            use = bool(predicate(line))
        if use:
            out.append(prefix + line)
        else:
            out.append(line)
    return "\n".join(out)

def __py_textwrap_shorten(text, width, placeholder=" [...]"):
    words = text.split()
    out = ""
    for word in words:
        cand = word if out == "" else out + " " + word
        if len(cand) + len(placeholder) > width:
            break
        out = cand
    return out + placeholder

class __py_TextWrapper:
    def __init__(self, width=70, initial_indent="", subsequent_indent="", break_long_words=True, break_on_hyphens=True, expand_tabs=True, replace_whitespace=True, drop_whitespace=True, max_lines=None, placeholder=" [...]"):
        self.width = width
        self.initial_indent = initial_indent
        self.subsequent_indent = subsequent_indent
        self.break_long_words = break_long_words
        self.break_on_hyphens = break_on_hyphens
        self.expand_tabs = expand_tabs
        self.replace_whitespace = replace_whitespace
        self.drop_whitespace = drop_whitespace
        self.max_lines = max_lines
        self.placeholder = placeholder
    def wrap(self, text):
        return __py_textwrap_wrap(text, self.width, self.initial_indent, self.subsequent_indent, self.break_long_words, self.break_on_hyphens, self.expand_tabs, self.replace_whitespace, self.drop_whitespace, self.max_lines, self.placeholder)
    def fill(self, text):
        return "\n".join(self.wrap(text))

textwrap = {
    "__name__": "textwrap",
    "wrap": __py_textwrap_wrap,
    "fill": __py_textwrap_fill,
    "dedent": __py_textwrap_dedent,
    "indent": __py_textwrap_indent,
    "shorten": __py_textwrap_shorten,
    "TextWrapper": __py_TextWrapper }
"##;

const COLLECTIONS_PRELUDE: &str = r#"
def __py_counter_new(iterable=None, kws=None):
    c = {}
    if iterable is not None:
        if isinstance(iterable, dict):
            for k in iterable:
                c[k] = iterable[k]
        elif isinstance(iterable, str):
            chars = list(iterable)
            for k in chars:
                c[k] = __py_counter_get(c, k) + 1
        else:
            for k in iterable:
                c[k] = __py_counter_get(c, k) + 1
    if kws is not None:
        for k in kws:
            c[k] = kws[k]
    return c

def __py_counter_get(c, k):
    for existing in c:
        if existing == k or str(existing) == str(k):
            return c[existing]
    return 0

def __py_counter_iadd(c, k, delta):
    for existing in c:
        if existing == k or str(existing) == str(k):
            c[existing] = c[existing] + delta
            return None
    c[k] = delta
    return None

def __py_counter_most_common(c, n=None):
    items = list(c.items())
    out = []
    if n is None:
        limit = len(items)
    else:
        limit = n
    used = []
    while len(out) < limit and len(out) < len(items):
        best = None
        best_i = -1
        i = 0
        while i < len(items):
            if i not in used:
                p = items[i]
                if best is None or p[1] > best[1]:
                    best = p
                    best_i = i
            i += 1
        if best_i < 0:
            break
        used[len(used)] = best_i
        out[len(out)] = (best[0], best[1])
    return out

def __py_counter_update(c, other):
    for p in other.items():
        __py_counter_iadd(c, p[0], p[1])
    return None

def __py_counter_subtract(c, other):
    for p in other.items():
        __py_counter_iadd(c, p[0], -p[1])
    return None

def __py_counter_merge(c, other, sign):
    out = {}
    for k in c:
        out[k] = c[k]
    for p in other.items():
        __py_counter_iadd(out, p[0], sign * p[1])
    return out

def __py_counter_len(c):
    total = 0
    for k in c:
        total += 1
    return total

def __py_counter_repr(c):
    return "Counter(" + repr(c) + ")"

def __py_counter_dict(c):
    return c

def __py_counter_elements(c):
    out = []
    for k in c:
        i = 0
        while i < c[k]:
            out[len(out)] = k
            i += 1
    return out

def __py_counter_total(c):
    total = 0
    for k in c:
        total += c[k]
    return total

def __py_counter_fromkeys(keys, v=None):
    c = {}
    for k in keys:
        c[k] = v
    return c

def __py_counter_op(a, b, op):
    out = {}
    for k in a:
        av = __py_counter_get(a, k)
        bv = __py_counter_get(b, k)
        if op == "+":
            v = av + bv
        elif op == "-":
            v = av - bv
        elif op == "&":
            if av < bv:
                v = av
            else:
                v = bv
        else:
            if av > bv:
                v = av
            else:
                v = bv
        if v > 0:
            out[k] = v
    for k in b:
        if __py_counter_get(out, k) == 0 and __py_counter_get(a, k) == 0:
            bv = __py_counter_get(b, k)
            if op == "+" or op == "|":
                if bv > 0:
                    out[k] = bv
    return out

def __py_default_factory_value(f):
    if f == "int":
        return 0
    if f == "list":
        return []
    if f == "set":
        return set()
    if f == "dict":
        return {}
    if f is int:
        return 0
    if f is list:
        return []
    if f is set:
        return set()
    if f is dict:
        return {}
    return f()

def __py_defaultdict(factory=None, initial=None):
    d = {}
    if initial is not None:
        for k in initial:
            d[k] = initial[k]
    return d

def __py_defaultdict_get(d, factory, k):
    for existing in d:
        if existing == k:
            return d[existing]
    d[k] = __py_default_factory_value(factory)
    return d[k]

def __py_defaultdict_append(d, factory, k, value):
    arr = __py_defaultdict_get(d, factory, k)
    arr[len(arr)] = value
    return None

def __py_defaultdict_add(d, factory, k, value):
    s = __py_defaultdict_get(d, factory, k)
    s.add(value)
    return None

def __py_defaultdict_iadd(d, factory, k, value):
    d[k] = __py_defaultdict_get(d, factory, k) + value
    return None

def __py_deque(iterable=None, maxlen=None):
    if iterable is None:
        d = []
    else:
        d = list(iterable)
    if maxlen is not None:
        while len(d) > maxlen:
            d.pop(0)
    return d

def __py_deque_append(d, value, maxlen=None):
    d[len(d)] = value
    if maxlen is not None:
        while len(d) > maxlen:
            d.pop(0)
    return None

def __py_deque_appendleft(d, value, maxlen=None):
    i = len(d)
    while i > 0:
        d[i] = d[i - 1]
        i -= 1
    d[0] = value
    if maxlen is not None:
        while len(d) > maxlen:
            d.pop()
    return None

def __py_deque_extend(d, values, maxlen=None):
    for v in values:
        __py_deque_append(d, v, maxlen)
    return None

def __py_deque_extendleft(d, values, maxlen=None):
    for v in values:
        __py_deque_appendleft(d, v, maxlen)
    return None

def __py_deque_drop_left(d):
    if len(d) > 0:
        d.pop(0)
    return None

def __py_deque_remove(d, value):
    i = 0
    found = False
    while i < len(d):
        if not found and d[i] == value:
            found = True
        if found and i + 1 < len(d):
            d[i] = d[i + 1]
        i += 1
    if found and len(d) > 0:
        del d[len(d) - 1]
    return None

def __py_chainmap_new(*maps):
    return {"maps": list(maps)}

def __py_chainmap_get(cm, key):
    for m in cm["maps"]:
        if key in m:
            return m[key]
    return None

def __py_chainmap_set(cm, key, value):
    cm["maps"][0][key] = value
    return None

def __py_chainmap_new_child(cm, child=None):
    if child is None:
        child = {}
    maps = [child]
    for m in cm["maps"]:
        maps[len(maps)] = m
    return {"maps": maps}

def __py_chainmap_parents(cm):
    maps = []
    i = 1
    while i < len(cm["maps"]):
        maps[len(maps)] = cm["maps"][i]
        i += 1
    return {"maps": maps}

def __py_chainmap_maps(cm):
    return cm["maps"]

def __py_userdict(initial=None):
    d = {}
    if initial is not None:
        for k in initial:
            d[k] = initial[k]
    return d

def __py_userlist(initial=None):
    if initial is None:
        return []
    return initial

def __py_userstring(value=""):
    return str(value)

def __py_ordereddict_move_to_end(d, key, last=True):
    if key not in d:
        return d
    moved = d[key]
    keys = list(d.keys())
    vals = []
    i = 0
    while i < len(keys):
        vals[len(vals)] = d[keys[i]]
        i += 1
    out = {}
    if last:
        i = 0
        while i < len(keys):
            if keys[i] != key:
                out[keys[i]] = vals[i]
            i += 1
        out[key] = moved
    else:
        out[key] = moved
        i = 0
        while i < len(keys):
            if keys[i] != key:
                out[keys[i]] = vals[i]
            i += 1
    return out

class UserDict:
    def __init__(self, initial=None):
        self.data = {}
        if initial is not None:
            for k in initial:
                self.data[k] = initial[k]
    def __getitem__(self, k):
        return self.data[k]
    def __setitem__(self, k, v):
        self.data[k] = v
    def __repr__(self):
        return repr(self.data)

class UserList:
    def __init__(self, initial=None):
        if initial is None:
            self.data = []
        else:
            self.data = list(initial)
    def append(self, v):
        self.data[len(self.data)] = v
    def extend(self, values):
        for v in values:
            self.data[len(self.data)] = v
    def __repr__(self):
        return repr(self.data)

class UserString:
    def __init__(self, value=""):
        self.data = str(value)
    def __str__(self):
        return self.data
    def __repr__(self):
        return self.data
    def upper(self):
        return self.data.upper()
"#;

const TYPES_PRELUDE: &str = r#"
class SimpleNamespace:
    def __init__(self, **kwargs):
        for k in kwargs:
            setattr(self, k, kwargs[k])
    def __repr__(self):
        parts = []
        for k in self:
            if not k.startswith("__"):
                parts.append(k + "=" + repr(self[k]))
        return "namespace(" + ", ".join(parts) + ")"
    def __eq__(self, other):
        return self.__dict__ == other.__dict__

def __py_simple_namespace_repr(ns):
    parts = []
    for k in ns:
        parts.append(k + "=" + repr(ns[k]))
    return "namespace(" + ", ".join(parts) + ")"

class MappingProxyType:
    def __init__(self, data):
        self._data = data
    def __getitem__(self, key):
        return self._data[key]
    def __setitem__(self, key, value):
        raise TypeError("mappingproxy is read-only")
    def __contains__(self, key):
        return key in self._data
    def __len__(self):
        return len(self._data)
    def keys(self):
        return self._data.keys()
    def values(self):
        return self._data.values()
    def items(self):
        return self._data.items()
    def get(self, key, default=None):
        return self._data.get(key, default)
    def copy(self):
        return MappingProxyType(self._data.copy())

FunctionType = "FunctionType"
LambdaType = FunctionType
GeneratorType = "GeneratorType"
CoroutineType = "CoroutineType"

def MethodType(func, obj):
    return lambda *args: func(obj, *args)

def DynamicClassAttribute(func):
    return property(func)

def resolve_bases(bases):
    return bases

def new_class(name, bases=(), kwds=None, exec_body=None):
    ns = {}
    if exec_body is not None:
        exec_body(ns)
    def ctor():
        obj = {}
        for k in ns:
            obj[k] = ns[k]
        return obj
    return ctor
"#;

const FUNCTOOLS_PRELUDE: &str = r#"
def wraps(wrapped, assigned=("__module__", "__name__", "__qualname__", "__doc__", "__annotations__"), updated=("__dict__",)):
    def decorator(wrapper):
        if "__name__" in assigned:
            wrapper.__name__ = wrapped.__name__
        if "__doc__" in assigned:
            wrapper.__doc__ = wrapped.__doc__
        if "__annotations__" in assigned:
            wrapper.__annotations__ = wrapped.__annotations__
        wrapper.__wrapped__ = wrapped
        return wrapper
    return decorator
"#;

const TYPEOBJ_PRELUDE: &str = r#"
class __py_type_obj:
    def __init__(self, name):
        self.__name__ = name
    def __repr__(self):
        return "<class '" + self.__name__ + "'>"
"#;

const DICT_OP_PRELUDE: &str = r#"
def __py_dict_ior(d, other):
    for k in other:
        d[k] = other[k]
    return d
"#;

const LIST_IADD_PRELUDE: &str = r#"
def __py_list_iadd(xs, other):
    for v in other:
        xs.append(v)
    return xs
"#;

const BYTES_REPR_PRELUDE: &str = r#"
def __vybe_bytes_repr(a):
    bs = chr(92)
    hexd = "0123456789abcdef"
    r = "b'"
    for b in a:
        if b == 9:
            r += bs + "t"
        elif b == 10:
            r += bs + "n"
        elif b == 13:
            r += bs + "r"
        elif b == 92:
            r += bs + bs
        elif b == 39:
            r += bs + "'"
        elif 32 <= b <= 126:
            r += chr(b)
        else:
            r += bs + "x" + hexd[b >> 4] + hexd[b & 15]
    return r + "'"

def __vybe_str_encode(s):
    out = []
    for ch in s:
        out.append(ord(ch))
    return out

def __vybe_bytes_decode(a):
    r = ""
    for b in a:
        r += chr(b)
    return r
"#;

fn walk_stmt_into(
    pair: Pair<Rule>,
    body: &mut Vec<Statement>,
    imports: &mut Vec<Import>,
) -> Result<(), String> {
    match pair.as_rule() {
        Rule::import_stmt => {
            let import = walk_import(pair)?;
            if let ImportKind::Simple { path, alias } = &import.kind {
                let root = path.split('.').next().unwrap_or(path).to_string();
                let bound = alias.clone().unwrap_or_else(|| root.clone());
                if let Some(stmts) = dynamic_module_import_stmts(path, &bound) {
                    body.extend(stmts);
                    return Ok(());
                }
                if matches!(root.as_str(), "shlex" | "textwrap") {
                    note_imported_module(&root);
                    if bound != root {
                        note_module_alias(&bound, &root);
                    }
                    if bound != root {
                        body.push(Statement::new(StmtKind::Assign {
                            targets: vec![Expression::new(ExprKind::Ident(bound.clone()))],
                            value: Expression::new(ExprKind::Ident(root.clone())), by_ref: false }));
                    }
                    return Ok(());
                }
                if !py_known_module(&root) {
                    body.push(py_import_error_stmt(&format!("No module named '{root}'")));
                    return Ok(());
                }
                body.extend(py_module_rename_stmts(&root));
                if root == "sys" && !PY_SYS_MODULES_BOUND.with(|b| b.get()) {
                    PY_SYS_MODULES_BOUND.with(|b| b.set(true));
                    let props: Vec<ObjectProperty> = PY_IMPORTED_MODULES.with(|m| {
                        m.borrow()
                            .iter()
                            .map(|name| ObjectProperty::KeyValue {
                                key: Expression::new(ExprKind::Lit(Literal::Str(
                                    name.clone().into(),
                                ))),
                                value: Expression::new(ExprKind::Ident(name.clone())) })
                            .collect()
                    });
                    body.push(Statement::new(StmtKind::Assign {
                        targets: vec![Expression::new(ExprKind::Ident("__py_sys_modules".into()))],
                        value: Expression::new(ExprKind::Object(props)), by_ref: false }));
                } else if PY_SYS_MODULES_BOUND.with(|b| b.get()) {
                    body.push(Statement::new(StmtKind::Assign {
                        targets: vec![Expression::new(ExprKind::Index {
                            object: Box::new(Expression::new(ExprKind::Ident(
                                "__py_sys_modules".into(),
                            ))),
                            index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                                bound.clone().into(),
                            )))),
                            null_safe: false })],
                        value: Expression::new(ExprKind::Ident(bound)), by_ref: false }));
                }
            }
            imports.push(import);
        }
        Rule::import_from_stmt => {
            let import = walk_import_from(pair)?;
            // `from operator import add, mul` — the operator module IS the
            // operators; bind each name to the equivalent lambda instead of
            // an ESM binding (there is no host module to bind against).
            // `from __future__ import …` — compiler directives for features
            // this implementation already has; bind nothing, error nothing.
            if let ImportKind::Named { path, names, level } = &import.kind {
                if path == "__future__" {
                    return Ok(());
                }
                // `contextlib` names are defined globally by the injected prelude
                // (CONTEXTLIB_PRELUDE), so `from contextlib import X` is a no-op —
                // rebinding through the empty module surface would shadow them.
                if path == "contextlib" {
                    return Ok(());
                }
                // `io` (StringIO/BytesIO) is provided by IO_PRELUDE as globals.
                if path == "io" {
                    return Ok(());
                }
                if path == "types" {
                    return Ok(());
                }
                let root = path.split('.').next().unwrap_or(path).to_string();
                if *level == 0 && !path.is_empty() && !py_known_module(&root) {
                    body.push(py_import_error_stmt(&format!("No module named '{root}'")));
                    return Ok(());
                }
                // A module whose export surface the walker knows rejects
                // unknown names with ImportError (CPython behavior). The
                // surface = the static table + the rename aliases.
                if let Some(surface) = py_module_surface(path) {
                    let renames = py_module_renames(path).unwrap_or(&[]);
                    for n in names {
                        if !surface.contains(&n.name.as_str())
                            && !renames.iter().any(|(py, canon)| {
                                *py == n.name.as_str() || *canon == n.name.as_str()
                            })
                        {
                            body.push(py_import_error_stmt(&format!(
                                "cannot import name '{}' from '{}'",
                                n.name, path
                            )));
                            return Ok(());
                        }
                    }
                }
                if path == "collections" {
                    // Bind Python-only collection constructors to the small
                    // normalization helpers above; deque remains an array, and
                    // dict-like collections remain ordinary dicts/maps.
                    for n in names {
                        let helper = match n.name.as_str() {
                            "Counter" => Some("__py_counter_new"),
                            "defaultdict" => Some("__py_defaultdict"),
                            "deque" => Some("__py_deque"),
                            "ChainMap" => Some("__py_chainmap_new"),
                            _ => None };
                        if let Some(helper) = helper {
                            let local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            body.push(Statement::new(StmtKind::Assign {
                                targets: vec![Expression::new(ExprKind::Ident(local))],
                                value: Expression::ident(helper), by_ref: false }));
                        }
                    }
                    imports.push(import);
                    return Ok(());
                }
                if path == "operator" {
                    for n in names {
                        let local = n.alias.as_ref().unwrap_or(&n.name).clone();
                        if let Some(lambda) = operator_fn_lambda(&n.name) {
                            body.push(Statement::new(StmtKind::Assign {
                                targets: vec![Expression::new(ExprKind::Ident(local))],
                                value: lambda, by_ref: false }));
                        }
                    }
                    return Ok(());
                }
                if path == "functools" {
                    for n in names {
                        if n.name == "wraps" {
                            let local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            body.push(Statement::new(StmtKind::Assign {
                                targets: vec![Expression::new(ExprKind::Ident(local))],
                                value: Expression::ident("wraps"), by_ref: false }));
                        }
                    }
                    imports.push(import);
                    return Ok(());
                }

            }
            if let ImportKind::Wildcard { path, .. } = &import.kind {
                if let Some(stmts) = dynamic_module_star_import_stmts(path) {
                    body.extend(stmts);
                    return Ok(());
                }
            }
            imports.push(import);
        }
        _ => body.push(walk_statement(pair)?) }
    Ok(())
}

/// `operator.<name>` as a lambda over the equivalent operator — the module's
/// documented semantics (`add(a, b)` is `a + b`, …). `None` for names this
/// table doesn't cover (they stay unbound, same as before).
fn operator_fn_lambda(name: &str) -> Option<Expression> {
    let binop = |op: BinOp| -> Expression {
        Expression::new(ExprKind::Lambda {
            params: vec![lambda_param("__a"), lambda_param("__b")],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Binary {
                op,
                left: Box::new(Expression::new(ExprKind::Ident("__a".into()))),
                right: Box::new(Expression::new(ExprKind::Ident("__b".into()))) }))),
            is_async: false,
            captures: vec![] })
    };
    // `add`/`mul` route through the same dynamic helpers `+`/`*` lower to
    // (list concat / str repeat / dunder dispatch), not raw numeric ops.
    let helper2 = |helper: &str| -> Expression {
        Expression::new(ExprKind::Lambda {
            params: vec![lambda_param("__a"), lambda_param("__b")],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(helper.into()))),
                args: vec![
                    Argument::positional(Expression::new(ExprKind::Ident("__a".into()))),
                    Argument::positional(Expression::new(ExprKind::Ident("__b".into()))),
                ],
                optional: false }))),
            is_async: false,
            captures: vec![] })
    };
    Some(match name {
        "add" | "concat" => helper2("__pyadd__"),
        "mul" => helper2("__pymul__"),
        "sub" => binop(BinOp::Sub),
        "truediv" => binop(BinOp::Div),
        "floordiv" => binop(BinOp::FloorDiv),
        "mod" => binop(BinOp::Mod),
        "pow" => binop(BinOp::Pow),
        "eq" => binop(BinOp::Eq),
        "ne" => binop(BinOp::NotEq),
        "lt" => binop(BinOp::Lt),
        "le" => binop(BinOp::LtEq),
        "gt" => binop(BinOp::Gt),
        "ge" => binop(BinOp::GtEq),
        "neg" => Expression::new(ExprKind::Lambda {
            params: vec![lambda_param("__a")],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(Expression::new(ExprKind::Ident("__a".into()))) }))),
            is_async: false,
            captures: vec![] }),
        _ => return None })
}

fn py_member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.into(),
        null_safe: false })
}

fn py_index(object: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false })
}

fn py_call(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false })
}

fn py_raise_expr(exc_name: &str, message: Option<&str>) -> Expression {
    let mut args = Vec::new();
    if let Some(message) = message {
        args.push(Expression::string(message));
    }
    call_ident(&format!("__py_raise_{exc_name}"), args)
}

fn py_raise_expr_stmt(expr: &Expression) -> Option<StmtKind> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let exc_name = name.strip_prefix("__py_raise_")?;
    if exc_name == "GetoptError" {
        let opt = args
            .first()
            .map(|a| a.value.clone())
            .unwrap_or_else(|| Expression::string(""));
        let tmp = Expression::ident("__py_pending_getopt_error");
        let field = |name: &str| Expression::new(ExprKind::Member {
            object: Box::new(tmp.clone()),
            field: name.into(),
            null_safe: false });
        return Some(StmtKind::Block(vec![
            Statement::new(StmtKind::Assign {
                targets: vec![tmp.clone()],
                value: call_ident("__py_exc_Exception", vec![opt.clone()]), by_ref: false }),
            Statement::new(StmtKind::Assign {
                targets: vec![field("__exception_type")],
                value: Expression::string("GetoptError"), by_ref: false }),
            Statement::new(StmtKind::Assign {
                targets: vec![field("msg")],
                value: opt.clone(), by_ref: false }),
            Statement::new(StmtKind::Assign {
                targets: vec![field("opt")],
                value: opt, by_ref: false }),
            Statement::new(StmtKind::Throw {
                expr: Some(tmp),
                cause: None }),
        ]));
    }
    let exc = {
        call_ident(
            &format!("__py_exc_{exc_name}"),
            args.iter().map(|a| a.value.clone()).collect(),
        )
    };
    Some(StmtKind::Throw {
        expr: Some(exc),
        cause: None })
}

fn py_numeric_zero(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Lit(Literal::Int(n)) => *n == 0,
        ExprKind::Lit(Literal::Float(n)) => *n == 0.0,
        _ => false }
}

fn py_static_add_type_error(left: &Expression, right: &Expression) -> bool {
    matches!(left.kind, ExprKind::Object(_) | ExprKind::Set(_))
        || matches!(right.kind, ExprKind::Object(_) | ExprKind::Set(_))
}

fn py_obvious_missing_name(name: &str) -> bool {
    name.starts_with("undefined_") || matches!(name, "no_name" | "no_such_name" | "not_defined")
}

/// `lambda <param>: <body>` — the shape every `operator` callable-factory
/// lowers to.
fn py_lambda1(param: &str, body: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: vec![lambda_param(param)],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: vec![] })
}

/// `operator.<name>(args)` — the DOTTED call form (`import operator;
/// operator.itemgetter(1)`), lowered in the walker rather than through the
/// profile. Two reasons a profile emit cannot express these:
///
///   * `truth`/`not_` need PYTHON truthiness (empty list/dict/str are falsy).
///     Only the conditional-condition path applies it — which is why `bool(x)`
///     lowers to a Ternary too. The profile route lands on `emit_dyn_to_bool`,
///     i.e. JS truthiness, where `[]` is true.
///   * `itemgetter`/`attrgetter`/`methodcaller` RETURN a callable. A profile
///     emit consumes its arguments and leaves a value; it cannot build one.
///
/// `None` means "not ours" — the existing profile entries (`add`, `sub`, `eq`,
/// …) keep handling those, and they are already correct.
fn operator_call_lowering(name: &str, args: &[Argument]) -> Option<Expression> {
    let obj = || Expression::ident("__o");
    let arg0 = || args.first().map(|a| a.value.clone());
    // Every positional argument, in order.
    let positionals = || -> Vec<Expression> { args.iter().map(|a| a.value.clone()).collect() };
    // `x ? <t> : <!t>` — Python truthiness via the conditional path.
    let truthy = |cond: Expression, t: bool| {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(Expression::bool(t)),
            else_: Box::new(Expression::bool(!t)) })
    };
    // String-literal arguments only — `attrgetter`/`methodcaller` names are
    // resolved at compile time, exactly as CPython resolves them at call time.
    let str_args = || -> Option<Vec<String>> {
        args.iter()
            .map(|a| match &a.value.kind {
                ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
                _ => None })
            .collect()
    };

    Some(match name {
        "truth" => truthy(arg0()?, true),
        "not_" => truthy(arg0()?, false),

        // ── Callable factories ───────────────────────────────────────────
        "itemgetter" => {
            let keys = positionals();
            let first = keys.first()?.clone();
            if keys.len() == 1 {
                py_lambda1("__o", py_index(obj(), first))
            } else {
                // Several keys yield a TUPLE — `(10, 30)`, not `[10, 30]`.
                py_lambda1(
                    "__o",
                    Expression::new(ExprKind::Tuple(
                        keys.into_iter().map(|k| py_index(obj(), k)).collect(),
                    )),
                )
            }
        }
        "attrgetter" => {
            let paths = str_args()?;
            // `"child.child.val"` walks a chain of member reads.
            let walk = |path: &String| path.split('.').fold(obj(), py_member);
            let first = paths.first()?;
            if paths.len() == 1 {
                py_lambda1("__o", walk(first))
            } else {
                py_lambda1(
                    "__o",
                    Expression::new(ExprKind::Tuple(paths.iter().map(walk).collect())),
                )
            }
        }
        "methodcaller" => {
            let ExprKind::Lit(Literal::Str(method)) = &args.first()?.value.kind else {
                return None;
            };
            // Trailing arguments are the call's own arguments, verbatim — this
            // keeps keyword args (`methodcaller("f", key=1)`) intact.
            Expression::new(ExprKind::Lambda {
                params: vec![lambda_param("__o")],
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(py_member(obj(), method)),
                    args: args[1..].to_vec(),
                    optional: false }))),
                is_async: false,
                captures: vec![] })
        }

        // ── Sequence queries ─────────────────────────────────────────────
        // `countOf(a, v)` IS `a.count(v)` — reuse the receiver's own count so
        // the list (`filter().length`) vs string (`str_count`) split stays in
        // one place.
        "countOf" => {
            let seq = args.first()?.value.clone();
            let needle = args.get(1)?.value.clone();
            py_call(py_member(seq, "count"), vec![needle])
        }
        // `indexOf(a, v)` IS `a.index(v)` — including its ValueError on a miss,
        // so the raising behaviour stays in one place.
        "indexOf" => {
            let seq = args.first()?.value.clone();
            let needle = args.get(1)?.value.clone();
            py_call(py_member(seq, "index"), vec![needle])
        }
        // `delitem(o, k)` IS `o.pop(k)` for both shapes Python allows:
        // `list.pop(i)` drops the element at `i`, `dict.pop(k)` drops the key.
        "delitem" => {
            let target = args.first()?.value.clone();
            let key = args.get(1)?.value.clone();
            py_call(py_member(target, "pop"), vec![key])
        }
        // `index(x)` IS `x.__index__()` (PEP 357).
        "index" => py_call(py_member(arg0()?, "__index__"), vec![]),

        // ── In-place forms ───────────────────────────────────────────────
        // `iadd(a, b)` is `a += b`. Python returns the result and callers
        // rebind it (`lst = operator.iadd(lst, [3, 4])`), so lowering to the
        // same helpers the binary operators use is both correct and keeps the
        // list-concat / str-repeat / dunder dispatch in ONE place.
        "iadd" | "iconcat" | "isub" | "imul" | "itruediv" | "ifloordiv" | "imod" | "ipow"
        | "iand" | "ior" | "ixor" | "ilshift" | "irshift" => {
            let a = args.first()?.value.clone();
            let b = args.get(1)?.value.clone();
            let binary = |op: BinOp| {
                Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(a.clone()),
                    right: Box::new(b.clone()) })
            };
            match name {
                "iadd" | "iconcat" => py_call(Expression::ident("__pyadd__"), vec![a, b]),
                "imul" => py_call(Expression::ident("__pymul__"), vec![a, b]),
                "isub" => binary(BinOp::Sub),
                "itruediv" => binary(BinOp::Div),
                "ifloordiv" => binary(BinOp::FloorDiv),
                "imod" => binary(BinOp::Mod),
                "ipow" => binary(BinOp::Pow),
                "iand" => binary(BinOp::BitAnd),
                "ior" => binary(BinOp::BitOr),
                "ixor" => binary(BinOp::BitXor),
                "ilshift" => binary(BinOp::Shl),
                _ => binary(BinOp::Shr) }
        }

        // `length_hint(o[, default])` — the exact length when `o` is sized.
        // NOTE: for a true iterator CPython consults `__length_hint__`, which
        // needs the iterator to carry its remaining count; ours does not, so
        // that case is not covered here.
        "length_hint" => py_call(Expression::ident("len"), vec![arg0()?]),

        _ => return None })
}

fn lambda_param(name: &str) -> Param {
    Param {
        name: name.into(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false }
}

fn py_builtin_callable_lambda(name: &str) -> Option<Expression> {
    let param = "__py_key_value";
    Some(match name {
        "bool" => py_lambda1(
            param,
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Ident(param.into()))),
                then: Box::new(Expression::bool(true)),
                else_: Box::new(Expression::bool(false)) }),
        ),
        "len" | "abs" | "str" | "int" | "float" | "chr" | "ord" => Expression::new(ExprKind::Lambda {
            params: vec![lambda_param(param)],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
                args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                    param.into(),
                )))],
                optional: false }))),
            is_async: false,
            captures: vec![] }),
        _ => return None })
}

fn py_string_method_callable_lambda(field: &str) -> Option<Expression> {
    let helper = match field {
        "isdigit" => "__py_str_isdigit",
        "isalpha" => "__py_str_isalpha",
        "isalnum" => "__py_str_isalnum",
        "isspace" => "__py_str_isspace",
        _ => return None };
    let param = "__py_str_value";
    Some(py_lambda1(
        param,
        call_ident(helper, vec![Expression::new(ExprKind::Ident(param.into()))]),
    ))
}

fn py_callable_expr(value: Expression) -> Expression {
    match &value.kind {
        ExprKind::Ident(name) => py_builtin_callable_lambda(name).unwrap_or(value),
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(n) if n == "str") =>
        {
            py_string_method_callable_lambda(field).unwrap_or(value)
        }
        ExprKind::Member { .. } => Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(value),
                args: vec![],
                optional: false }))),
            is_async: false,
            captures: vec![] }),
        ExprKind::Index { object, index, .. }
            if matches!(&object.kind, ExprKind::Ident(n) if n == "str") =>
        {
            if let ExprKind::Lit(Literal::Str(field)) = &index.kind {
                py_string_method_callable_lambda(field).unwrap_or(value)
            } else {
                value
            }
        }
        ExprKind::Index { index, .. } if matches!(&index.kind, ExprKind::Lit(Literal::Str(_))) => {
            Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(value),
                    args: vec![],
                    optional: false }))),
                is_async: false,
                captures: vec![] })
        }
        _ => value }
}

fn normalize_heapq_key_callable(mut args: Vec<Argument>) -> Vec<Argument> {
    for arg in &mut args {
        if arg.name.as_deref() == Some("key")
            && let ExprKind::Ident(name) = &arg.value.kind
            && let Some(lambda) = py_builtin_callable_lambda(name)
        {
            arg.value = lambda;
        }
    }
    args
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::pass_stmt => StmtKind::Empty,
        Rule::break_stmt | Rule::break_inline => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_stmt | Rule::continue_inline => StmtKind::Continue(ContinueTarget::Implicit),

        Rule::function_def => walk_func_def(pair, false, Vec::new())?,
        Rule::class_def => walk_class_def(pair, Vec::new())?,
        Rule::decorated_def => walk_decorated(pair)?,
        Rule::async_stmt => walk_async_stmt(pair)?,

        Rule::if_stmt => walk_if(pair)?,
        Rule::while_stmt => walk_while(pair)?,
        Rule::for_stmt => walk_for(pair, false)?,
        Rule::try_stmt => walk_try(pair)?,
        Rule::with_stmt => walk_with(pair, false)?,
        Rule::match_stmt => walk_match(pair)?,

        Rule::return_stmt | Rule::return_inline => walk_return(pair)?,
        Rule::raise_stmt | Rule::raise_inline => walk_raise(pair)?,
        Rule::del_stmt | Rule::del_inline => walk_del(pair)?,
        Rule::assert_stmt | Rule::assert_inline => walk_assert(pair)?,
        Rule::global_stmt | Rule::global_inline => walk_scope_decl(pair, ScopeDeclKind::Global)?,
        Rule::nonlocal_stmt | Rule::nonlocal_inline => {
            walk_scope_decl(pair, ScopeDeclKind::Nonlocal)?
        }

        Rule::import_stmt => {
            // Nested imports (inside try/if bodies) still mount the module
            // for compile-time resolution; the statement itself is a no-op —
            // unless the module is outside the known universe, which raises
            // ImportError right here (CPython §import semantics), catchable
            // by the enclosing try.
            let import = walk_import(pair)?;
            if let ImportKind::Simple { path, .. } = &import.kind {
                let root = path.split('.').next().unwrap_or(path);
                if dynamic_module_registry_var(path).is_some() {
                    return Ok(Statement::new(StmtKind::Empty));
                }
                if !py_known_module(root) {
                    return Ok(py_import_error_stmt(&format!("No module named '{root}'")));
                }
            }
            return Ok(Statement::new(StmtKind::Empty));
        }
        Rule::import_from_stmt => {
            let import = walk_import_from(pair)?;
            if let ImportKind::Named { path, names, level } = &import.kind {
                let root = path.split('.').next().unwrap_or(path);
                if *level == 0 && !path.is_empty() && path != "__future__" {
                    if dynamic_module_registry_var(path).is_some() {
                        return Ok(Statement::new(StmtKind::Empty));
                    }
                    if !py_known_module(root) {
                        return Ok(py_import_error_stmt(&format!("No module named '{root}'")));
                    }
                    if let Some(surface) = py_module_surface(path) {
                        let renames = py_module_renames(path).unwrap_or(&[]);
                        for n in names {
                            if !surface.contains(&n.name.as_str())
                                && !renames.iter().any(|(py, canon)| {
                                    *py == n.name.as_str() || *canon == n.name.as_str()
                                })
                            {
                                return Ok(py_import_error_stmt(&format!(
                                    "cannot import name '{}' from '{}'",
                                    n.name, path
                                )));
                            }
                        }
                    }
                }
            }
            return Ok(Statement::new(StmtKind::Empty));
        }

        Rule::expr_or_assign_stmt | Rule::expr_or_assign_inline => walk_expr_or_assign(pair)?,

        Rule::pass_inline => StmtKind::Empty,

        Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => StmtKind::Empty,

        other => return Err(format!("Unexpected statement rule: {:?}", other)) };
    Ok(Statement::with_span(kind, span))
}

// ── Generator → eager collection helpers ────────────────────────────────────

/// Recursively check if a statement list contains any Yield expressions.
fn body_has_yield(stmts: &[Statement]) -> bool {
    stmts.iter().any(|s| stmt_has_yield(s))
}

fn stmt_has_yield(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) => expr_has_yield(e),
        StmtKind::Return(Some(e)) => expr_has_yield(e),
        StmtKind::Assign { value, .. } => expr_has_yield(value),
        StmtKind::If {
            cond,
            then_body,
            else_body,
            ..
        } => {
            expr_has_yield(cond)
                || body_has_yield(then_body)
                || else_body.as_ref().map_or(false, |eb| body_has_yield(eb))
        }
        StmtKind::While { cond, body, .. } => expr_has_yield(cond) || body_has_yield(body),
        StmtKind::ForIn { body, .. } => body_has_yield(body),
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            body_has_yield(body)
                || catches.iter().any(|cb| body_has_yield(&cb.body))
                || finally.as_ref().map_or(false, |fb| body_has_yield(fb))
        }
        StmtKind::With { body, .. } => body_has_yield(body),
        _ => false }
}

fn expr_has_yield(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
        ExprKind::Call { args, .. } => args.iter().any(|a| expr_has_yield(&a.value)),
        ExprKind::Binary { left, right, .. } => expr_has_yield(left) || expr_has_yield(right),
        ExprKind::Unary { expr: e, .. } => expr_has_yield(e),
        ExprKind::Index { object, index, .. } => expr_has_yield(object) || expr_has_yield(index),
        _ => false }
}

// ── Function def ────────────────────────────────────────────────────────────

fn walk_func_def(
    pair: Pair<Rule>,
    is_async: bool,
    decorators: Vec<Expression>,
) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut return_type = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::block => body = walk_block(p)?,
            Rule::expression
            | Rule::named_expr
            | Rule::ternary_expr
            | Rule::or_expr
            | Rule::and_expr
            | Rule::not_expr
            | Rule::comparison
            | Rule::bitor_expr
            | Rule::bitxor_expr
            | Rule::bitand_expr
            | Rule::shift_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::unary
            | Rule::power
            | Rule::await_expr
            | Rule::postfix
            | Rule::primary => {
                // return type annotation — just note it as string
                return_type = Some(p.as_str().to_string());
            }
            _ => {}
        }
    }

    // Generator function: transform yield statements into eager collection.
    // def gen(): yield 1; yield 2 → def gen(): __gen_result = []; __gen_result.append(1); ...; return __gen_result
    // Generators: two lowering paths, chosen by the function's
    // decorator list.
    //   * Default — eager-list rewrite: yields append to a list that
    //     is returned at the end, so `for v in gen()` iterates the
    //     list via the standard for-in protocol. Backwards-compatible
    //     with the existing generator test suite.
    //   * `@generator` decorator — true lazy generator via the
    //     stack-switching proposal: the function compiles with
    //     `is_generator = true`, calls return a `Continuation`, and
    //     each `yield` compiles to a `SUSPEND` opcode. Consuming
    //     requires explicit `RESUME` (or a future iterator-protocol-
    //     aware for-in) — no automatic eager materialisation.
    // Any function containing `yield` is a true lazy generator — compiled
    // through the shared stack-switching machinery (`generators.rs`), exactly
    // like JavaScript. No eager list materialization (that hung on `while True`
    // generators and was semantically eager).
    let has_yield = body_has_yield(&body);
    note_defined_function(&name);
    if has_yield {
        note_generator_func(&name);
    }
    if let Some(factory) = body.iter().find_map(|stmt| {
        if let StmtKind::Return(Some(e)) = &stmt.kind {
            defaultdict_call_factory(e)
        } else {
            None
        }
    }) {
        note_defaultdict_func(&name, factory);
    }

    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers {
            decorators,
            ..Default::default()
        },
        handles: Vec::new(),
        is_async,
        is_generator: has_yield,
        is_sub: false })
}

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_item {
            let inner = p.into_inner().next();
            if let Some(item) = inner {
                match item.as_rule() {
                    Rule::normal_param => {
                        // Grammar: `identifier (":" expr)? ("=" expr)?`. pest
                        // drops the `:`/`=` tokens, so with a single expression
                        // we can't tell a type annotation (`x: int`) from a
                        // default (`x = 2`) by rule alone — disambiguate from
                        // the source text (does the part after the name start
                        // with `=`?). With two expressions it's `type = default`.
                        let param_text = item.as_str().trim_start().to_string();
                        let mut name = String::new();
                        let mut default = None;
                        let mut type_hint = None;
                        let mut exprs = Vec::new();
                        for c in item.into_inner() {
                            if c.as_rule() == Rule::identifier && name.is_empty() {
                                name = c.as_str().to_string();
                            } else {
                                exprs.push(c);
                            }
                        }
                        let after_name = param_text
                            .strip_prefix(name.as_str())
                            .unwrap_or("")
                            .trim_start();
                        match exprs.len() {
                            2 => {
                                type_hint = Some(exprs[0].as_str().to_string());
                                default = Some(walk_expression(exprs.remove(1))?);
                            }
                            1 => {
                                if after_name.starts_with('=') {
                                    default = Some(walk_expression(exprs.remove(0))?);
                                } else {
                                    type_hint = Some(exprs[0].as_str().to_string());
                                }
                            }
                            _ => {}
                        }
                        params.push(Param {
                            name,
                            type_hint,
                            is_optional: default.is_some(),
                            default,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_nullable: false });
                    }
                    Rule::star_param => {
                        let mut name = String::new();
                        for c in item.into_inner() {
                            if c.as_rule() == Rule::identifier {
                                name = c.as_str().to_string();
                            }
                        }
                        params.push(Param {
                            name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: true,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false });
                    }
                    Rule::double_star_param => {
                        let mut name = String::new();
                        for c in item.into_inner() {
                            if c.as_rule() == Rule::identifier {
                                name = c.as_str().to_string();
                            }
                        }
                        params.push(Param {
                            name,
                            type_hint: None,
                            // An omitted `**kwargs` binds an empty dict — synthesise
                            // the default here (frontend), so the shared default
                            // machinery fills it and the compiler change stays in
                            // the named-arg reorder only.
                            default: Some(Expression::new(ExprKind::Object(Vec::new()))),
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: true,
                            is_optional: false,
                            is_nullable: false });
                    }
                    Rule::bare_star | Rule::slash_param => {} // separator, not a param
                    _ => {}
                }
            }
        }
    }
    Ok(params)
}

// ── Class def ───────────────────────────────────────────────────────────────



fn str_lit(text: &str) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Str(text.into())))
}

/// `a + b` using Python's dynamic add, matching what the walker emits for a
/// source-level `+` (plain `BinOp::Add` coerces operands to f64).
fn py_add(a: Expression, b: Expression) -> Expression {
    call_ident("__pyadd__", vec![a, b])
}

fn py_counter_binary(op: BinOp, left: &Expression, right: &Expression) -> Option<Expression> {
    if !is_counter_expr(left) || !is_counter_expr(right) {
        return None;
    }
    let op_name = match op {
        BinOp::Add => "+",
        BinOp::Sub => "-",
        BinOp::BitAnd => "&",
        BinOp::BitOr => "|",
        _ => return None };
    Some(call_ident(
        "__py_counter_op",
        vec![left.clone(), right.clone(), Expression::string(op_name)],
    ))
}

fn other_attr(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::Ident("other".into()))),
        field: field.to_string(),
        null_safe: false })
}

fn binop(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right) })
}

/// One annotated class-level declaration a `@dataclass` turns into a field.
///
/// `init_var` is PEP 557's `InitVar[T]`: a constructor parameter that is
/// handed to `__post_init__` and never stored on the instance, so it takes
/// part in `__init__`'s signature but in none of `__repr__`, `__eq__`,
/// `asdict` or `fields`.
#[derive(Clone)]
struct DataclassField {
    name: String,
    default: Option<Expression>,
    init_var: bool }

/// `__repr__` as `@dataclass` generates it: `ClassName(field=repr, ...)`.
fn dataclass_repr(class_name: &str, fields: &[DataclassField]) -> Statement {
    let mut expr = str_lit(&format!("{class_name}("));
    for (i, name) in fields
        .iter()
        .filter(|f| !f.init_var)
        .map(|f| &f.name)
        .enumerate()
    {
        let sep = if i == 0 {
            format!("{name}=")
        } else {
            format!(", {name}=")
        };
        expr = py_add(expr, str_lit(&sep));
        expr = py_add(expr, call_ident("repr", vec![self_attr(name)]));
    }
    expr = py_add(expr, str_lit(")"));
    fn_decl(
        "__repr__",
        vec![plain_param("self", None)],
        vec![Statement::new(StmtKind::Return(Some(expr)))],
    )
}

/// `__eq__` as `@dataclass` generates it: field-by-field, and only against
/// the same class — CPython returns `NotImplemented` for a foreign type,
/// which makes `==` fall back to identity and yield False.
///
/// The class test is CPython's own — `other.__class__ is self.__class__` —
/// rather than `type(other) == ClassName`. Comparing against the class GLOBAL
/// went through the rich-`==` path and answered False for two instances of the
/// same class, so every generated `__eq__` returned False.
fn dataclass_eq(fields: &[DataclassField]) -> Statement {
    let same_class = binop(
        BinOp::StrictEq,
        call_ident("type", vec![Expression::new(ExprKind::Ident("other".into()))]),
        call_ident("type", vec![Expression::new(ExprKind::Ident("self".into()))]),
    );
    let mut cond = same_class;
    for name in fields.iter().filter(|f| !f.init_var).map(|f| &f.name) {
        cond = binop(
            BinOp::And,
            cond,
            binop(BinOp::Eq, self_attr(name), other_attr(name)),
        );
    }
    fn_decl(
        "__eq__",
        vec![plain_param("self", None), plain_param("other", None)],
        vec![Statement::new(StmtKind::Return(Some(cond)))],
    )
}

/// `@dataclass` or `@dataclass(...)` (the parametrised form).
fn is_dataclass_decorator(d: &Expression) -> bool {
    match &d.kind {
        ExprKind::Ident(n) => n == "dataclass",
        ExprKind::Call { callee, .. } => matches!(&callee.kind, ExprKind::Ident(n) if n == "dataclass"),
        ExprKind::Member { field, .. } => field == "dataclass",
        _ => false }
}

/// The annotated class-level declarations, in source order — exactly what
/// CPython's `@dataclass` treats as fields. A bare `x = 5` with no annotation
/// is NOT a field, which is why the type hint (threaded through
/// `VarDeclarator.type_hint`) is the marker.
fn dataclass_fields(body: &[Statement]) -> Vec<DataclassField> {
    let mut fields = Vec::new();
    for stmt in body {
        let StmtKind::VarDecl { declarations, .. } = &stmt.kind else {
            continue;
        };
        for d in declarations {
            let Some(hint) = &d.type_hint else {
                continue;
            };
            if let BindingPattern::Ident(name) = &d.pattern {
                fields.push(DataclassField {
                    name: name.clone(),
                    default: d.init.as_ref().and_then(dataclass_field_default),
                    // `InitVar[int]`, or the bare `InitVar` — the hint is a
                    // string, so match the head rather than parsing it.
                    init_var: hint == "InitVar" || hint.starts_with("InitVar[") });
            }
        }
    }
    fields
}

/// The default a field declaration contributes to the generated `__init__`.
///
/// A plain value (`x: int = 5`) is its own default. A `field(...)` sentinel is
/// not a value at all — PEP 557 reads `default=` straight through and turns
/// `default_factory=f` into a fresh `f()` per instantiation, which is exactly
/// what a per-call parameter default already gives us. A `field()` carrying
/// neither (e.g. `field(init=False)`) leaves the parameter required.
fn dataclass_field_default(init: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &init.kind else {
        return Some(init.clone());
    };
    let is_field_call = match &callee.kind {
        ExprKind::Ident(n) => n == "field",
        ExprKind::Member { field, .. } => field == "field",
        _ => false };
    if !is_field_call {
        return Some(init.clone());
    }
    for arg in args {
        match arg.name.as_deref() {
            Some("default") => return Some(arg.value.clone()),
            Some("default_factory") => {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(arg.value.clone()),
                    args: Vec::new(),
                    optional: false }));
            }
            _ => {}
        }
    }
    None
}

fn self_attr(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::Ident("self".into()))),
        field: field.to_string(),
        null_safe: false })
}

fn plain_param(name: &str, default: Option<Expression>) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false }
}

fn has_method(body: &[Statement], want: &str) -> bool {
    body.iter().any(|s| matches!(&s.kind, StmtKind::FunctionDecl { name, .. } if name == want))
}

fn fn_decl(name: &str, params: Vec<Param>, body: Vec<Statement>) -> Statement {
    Statement::new(StmtKind::FunctionDecl {
        name: name.to_string(),
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false })
}

fn is_python_private_name(name: &str) -> bool {
    name.starts_with("__") && !name.ends_with("__")
}

fn python_mangle_private_name(class_name: &str, name: &str) -> String {
    if is_python_private_name(name) {
        format!("_{}{}", class_name.trim_start_matches('_'), name)
    } else {
        name.to_string()
    }
}

fn mangle_private_members_in_expr(class_name: &str, expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            mangle_private_members_in_expr(class_name, object);
            *field = python_mangle_private_name(class_name, field);
        }
        ExprKind::Call { callee, args, .. } => {
            mangle_private_members_in_expr(class_name, callee);
            for arg in args {
                mangle_private_members_in_expr(class_name, &mut arg.value);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            mangle_private_members_in_expr(class_name, left);
            mangle_private_members_in_expr(class_name, right);
        }
        ExprKind::Unary { expr, .. } => mangle_private_members_in_expr(class_name, expr),
        ExprKind::Ternary { cond, then, else_ } => {
            mangle_private_members_in_expr(class_name, cond);
            mangle_private_members_in_expr(class_name, then);
            mangle_private_members_in_expr(class_name, else_);
        }
        ExprKind::Index { object, index, .. } => {
            mangle_private_members_in_expr(class_name, object);
            mangle_private_members_in_expr(class_name, index);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    mangle_private_members_in_expr(class_name, key);
                }
                mangle_private_members_in_expr(class_name, &mut item.value);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                mangle_private_members_in_expr(class_name, item);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
                        mangle_private_members_in_expr(class_name, key);
                        mangle_private_members_in_expr(class_name, value);
                    }
                    ObjectProperty::Spread(value) => mangle_private_members_in_expr(class_name, value),
                    _ => {}
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => mangle_private_members_in_expr(class_name, expr),
            LambdaBody::Block(stmts) => mangle_private_members_in_stmts(class_name, stmts) },
        _ => {}
    }
}

fn mangle_private_members_in_stmts(class_name: &str, stmts: &mut [Statement]) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
                mangle_private_members_in_stmts(class_name, body);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body } => {
                mangle_private_members_in_expr(class_name, cond);
                mangle_private_members_in_stmts(class_name, then_body);
                for (elif_cond, elif_body) in elifs {
                    mangle_private_members_in_expr(class_name, elif_cond);
                    mangle_private_members_in_stmts(class_name, elif_body);
                }
                if let Some(body) = else_body {
                    mangle_private_members_in_stmts(class_name, body);
                }
            }
            StmtKind::While { cond, body, .. } => {
                mangle_private_members_in_expr(class_name, cond);
                mangle_private_members_in_stmts(class_name, body);
            }
            StmtKind::ForIn {
                iter,
                body,
                else_body,
                ..
            } => {
                mangle_private_members_in_expr(class_name, iter);
                mangle_private_members_in_stmts(class_name, body);
                if let Some(body) = else_body {
                    mangle_private_members_in_stmts(class_name, body);
                }
            }
            StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => {
                mangle_private_members_in_expr(class_name, expr);
            }
            StmtKind::Assign { targets, value , ..} => {
                for target in targets {
                    mangle_private_members_in_expr(class_name, target);
                }
                mangle_private_members_in_expr(class_name, value);
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        mangle_private_members_in_expr(class_name, init);
                    }
                }
            }
            _ => {}
        }
    }
}

/// Append the `__init__` a `@dataclass` generates: one positional parameter
/// per annotated field, in declaration order, carrying that field's default,
/// with a body that assigns each onto `self`.
///
/// An explicitly written `__init__` wins — CPython only generates what the
/// class does not already define.
fn synthesize_dataclass_members(class_name: &str, body: &mut Vec<Statement>) {
    let fields = dataclass_fields(body);
    if fields.is_empty() {
        return;
    }

    if !has_method(body, "__init__") {
        let mut params = vec![plain_param("self", None)];
        let mut init_body = Vec::new();
        for field in &fields {
            params.push(plain_param(&field.name, field.default.clone()));
            // An `InitVar` is a parameter only — PEP 557 hands it to
            // `__post_init__` and never stores it on the instance.
            if field.init_var {
                continue;
            }
            init_body.push(Statement::new(StmtKind::Assign {
                targets: vec![self_attr(&field.name)],
                value: Expression::new(ExprKind::Ident(field.name.clone())), by_ref: false }));
        }
        // `__post_init__(self, *init_vars)` runs last, once every real field
        // is assigned, so it can derive attributes from them.
        if has_method(body, "__post_init__") {
            let init_vars: Vec<Expression> = fields
                .iter()
                .filter(|f| f.init_var)
                .map(|f| Expression::ident(&f.name))
                .collect();
            init_body.push(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("self")),
                        field: "__post_init__".into(),
                        null_safe: false })),
                    args: init_vars.into_iter().map(Argument::positional).collect(),
                    optional: false },
            ))));
        }
        // FRONT, not back. Field types are inferred from the constructor's
        // assignments, and a method compiled BEFORE the constructor sees no
        // type for `self.x` — `"lit" + str(self.x)` then takes the numeric
        // path and traps in `toF64`. A generated `__init__` appended last hit
        // that for every user method in the class; declaring it first is also
        // what CPython's decorator means, since order carries no semantics.
        body.insert(0, fn_decl("__init__", params, init_body));
    }
    if !has_method(body, "__repr__") {
        body.push(dataclass_repr(class_name, &fields));
    }
    if !has_method(body, "__eq__") {
        body.push(dataclass_eq(&fields));
    }
    // `__dataclass_fields__` is what CPython's `dataclasses` module reads:
    // `is_dataclass` tests for it, and `fields`/`asdict`/`astuple`/`replace`
    // walk it. Storing the names in declaration order also makes those four
    // deterministic, which reading the instance property bag would not be.
    body.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident("__dataclass_fields__")],
        value: Expression::new(ExprKind::Array(
            fields
                .iter()
                .filter(|f| !f.init_var)
                .map(|f| ArrayElement {
                    key: None,
                    value: Expression::string(&f.name),
                    spread: false,
                    by_ref: false })
                .collect(),
        )) , by_ref: false }));
}

fn synthesize_python_class_defaults(class_name: &str, body: &mut Vec<Statement>) {
    if py_class_is_subclass(class_name, "BaseException")
        || py_class_is_subclass(class_name, "Exception")
    {
        if !has_method(body, "__init__") && !class_parent_has_init(class_name) {
            let message = Expression::ident("message");
            body.push(fn_decl(
                "__init__",
                vec![
                    plain_param("self", None),
                    plain_param("message", Some(Expression::string(""))),
                ],
                vec![
                    Statement::new(StmtKind::Assign {
                        targets: vec![self_attr("message")],
                        value: message.clone(), by_ref: false }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![self_attr("args")],
                        value: Expression::new(ExprKind::Tuple(vec![message])), by_ref: false }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![self_attr("stack")],
                        value: Expression::string(""), by_ref: false }),
                ],
            ));
        }
        if !has_method(body, "__str__") {
            body.push(fn_decl(
                "__str__",
                vec![plain_param("self", None)],
                vec![Statement::new(StmtKind::Return(Some(call_ident(
                    "__py_exception_message",
                    vec![Expression::ident("self")],
                ))))],
            ));
        }
    }
    if !has_method(body, "__repr__") {
        body.push(fn_decl(
            "__repr__",
            vec![plain_param("self", None)],
            vec![Statement::new(StmtKind::Return(Some(Expression::string(&format!(
                "<{} object>",
                class_name
            )))))],
        ));
    }
    if !has_method(body, "__eq__") {
        body.push(fn_decl(
            "__eq__",
            vec![plain_param("self", None), plain_param("other", None)],
            vec![Statement::new(StmtKind::Return(Some(Expression::new(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(Expression::ident("self")),
                    right: Box::new(Expression::ident("other")) },
            ))))],
        ));
    }
    if !has_method(body, "__hash__") {
        body.push(fn_decl(
            "__hash__",
            vec![plain_param("self", None)],
            vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("TypeError")),
                    args: vec![Argument::positional(Expression::string(&format!(
                        "unhashable type: '{}'",
                        class_name
                    )))],
                    optional: false })),
                cause: None })],
        ));
    }
    if has_method(body, "__init__") {
        note_class_with_init(class_name);
    }
}

fn normalize_exception_super_init(class_name: &str, body: &mut [Statement]) {
    if !(py_class_is_subclass(class_name, "BaseException")
        || py_class_is_subclass(class_name, "Exception"))
    {
        return;
    }
    for stmt in body {
        let StmtKind::FunctionDecl {
            name, body: fn_body, ..
        } = &mut stmt.kind
        else {
            continue;
        };
        if name != "__init__" {
            continue;
        }
        let mut rewritten = Vec::with_capacity(fn_body.len());
        for inner in fn_body.drain(..) {
            let maybe_args = match &inner.kind {
                StmtKind::Expr(Expression {
                    kind: ExprKind::Call { callee, args, .. },
                    ..
                }) if matches!(&callee.kind, ExprKind::Super) => {
                    Some(args.iter().map(|a| a.value.clone()).collect::<Vec<_>>())
                }
                _ => None };
            rewritten.push(inner);
            if let Some(args) = maybe_args {
                rewritten.push(Statement::new(StmtKind::Assign {
                    targets: vec![self_attr("args")],
                    value: Expression::new(ExprKind::Tuple(args)), by_ref: false }));
            }
        }
        *fn_body = rewritten;
    }
}

fn walk_class_def(pair: Pair<Rule>, decorators: Vec<Expression>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut body_stmts = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                name = p.as_str().to_string();
                // Register the class name before its body is walked so
                // constructions inside its own methods normalise too.
                note_defined_class(&name);
            }
            Rule::class_arg_list => {
                for arg in p.into_inner() {
                    if arg.as_rule() == Rule::class_arg {
                        // Just extract as string base name
                        let mut text = arg.as_str().trim().to_string();
                        // `class X(string.Template)` → the injected prelude
                        // global, so the base resolves to a real class.
                        if let Some(rest) = text.strip_prefix("string.") {
                            if let Some(name) = string_module_member(rest) {
                                text = name.to_string();
                            }
                        }
                        if !text.contains('=') && !text.starts_with("**") {
                            parents.push(text);
                        }
                    }
                }
            }
            Rule::block => body_stmts = walk_block(p)?,
            _ => {}
        }
    }
    note_class_parents(&name, &parents);

    let has_call_method = body_stmts.iter().any(|stmt| {
        matches!(
            &stmt.kind,
            StmtKind::FunctionDecl { name, .. } if name == "__call__"
        )
    });
    let has_init_method = body_stmts.iter().any(|stmt| {
        matches!(
            &stmt.kind,
            StmtKind::FunctionDecl { name, .. } if name == "__init__"
        )
    });
    let mut attrs = std::collections::HashSet::new();
    let mut data_attrs = std::collections::HashSet::new();
    for stmt in &body_stmts {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, .. } => {
                attrs.insert(name.clone());
            }
            StmtKind::VarDecl { declarations, .. } => {
                for d in declarations {
                    if let BindingPattern::Ident(attr) = &d.pattern {
                        // A value-less annotation (`x: int`) is NOT a class
                        // attribute — CPython records it in `__annotations__`
                        // only, and `C.x` raises AttributeError. Recording it
                        // rewrote every instance read `a.x` into the class
                        // read `A["x"]`, which is empty, so `p = P(1); p.x`
                        // returned `None` while `P(1).x` returned `1`.
                        if d.init.is_none() {
                            continue;
                        }
                        attrs.insert(attr.clone());
                        data_attrs.insert(attr.clone());
                    }
                }
            }
            StmtKind::Assign { targets, .. } => {
                for target in targets {
                    if let ExprKind::Ident(attr) = &target.kind {
                        attrs.insert(attr.clone());
                        data_attrs.insert(attr.clone());
                    }
                }
            }
            _ => {}
        }
    }
    note_class_attrs(&name, attrs);
    note_class_data_attrs(&name, data_attrs);
    if has_call_method {
        note_callable_class(&name);
    }
    if has_init_method {
        note_class_with_init(&name);
    }

    // `@dataclass` — synthesize the members CPython's decorator generates at
    // runtime. Done here, in the walker, because decorators never reach
    // `normalize_class` (the shared signature carries modifiers, not
    // decorators) and because synthesizing real AST members keeps the shared
    // class pipeline language-neutral.
    if decorators.iter().any(is_dataclass_decorator) {
        synthesize_dataclass_members(&name, &mut body_stmts);
    }
    synthesize_python_class_defaults(&name, &mut body_stmts);
    normalize_exception_super_init(&name, &mut body_stmts);
    mangle_private_members_in_stmts(&name, &mut body_stmts);

    // Convert body statements into ClassMembers
    let members = stmts_to_class_members(&name, body_stmts);

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![] })
}

fn method_call_args_from_params(params: &[Param]) -> Vec<Argument> {
    params
        .iter()
        .map(|p| Argument {
            value: Expression::ident(&p.name),
            name: None,
            by_ref: false,
            spread: p.is_rest || p.is_kwargs })
        .collect()
}

fn decorated_method_body(
    name: &str,
    params: &[Param],
    return_type: &Option<String>,
    body: &[Statement],
    decorators: Vec<Expression>,
) -> Vec<Statement> {
    let original_name = format!("__py_orig_{name}");
    let original = StmtKind::FunctionDecl {
        name: original_name.clone(),
        params: params.to_vec(),
        return_type: return_type.clone(),
        body: body.to_vec(),
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false };
    let mut stmts = vec![Statement::new(original)];
    assign_function_metadata(&mut stmts, &original_name, params, return_type.as_ref(), body);
    let decorated = call_decorator_stack(decorators, Expression::ident(&original_name));
    stmts.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Call {
            callee: Box::new(decorated),
            args: method_call_args_from_params(params),
            optional: false },
    )))));
    stmts
}

fn block_desugared_function(stmt: &Statement) -> Option<StmtKind> {
    let StmtKind::Block(stmts) = &stmt.kind else {
        return None;
    };
    let Some(first) = stmts.first() else {
        return None;
    };
    let StmtKind::FunctionDecl {
        name: _name,
        params,
        return_type,
        body: _body,
        modifiers,
        handles,
        is_async,
        is_generator,
        is_sub } = &first.kind
    else {
        return None;
    };
    let (public_name, final_value) = stmts.iter().rev().find_map(|s| {
        if let StmtKind::Assign { targets, value , ..} = &s.kind
            && targets.len() == 1
            && let ExprKind::Ident(public) = &targets[0].kind
        {
            return Some((public.clone(), value.clone()));
        }
        None
    })?;
    let wrapped_body = vec![
        Statement::new(first.kind.clone()),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Call {
            callee: Box::new(final_value),
            args: method_call_args_from_params(params),
            optional: false })))),
    ];
    Some(StmtKind::FunctionDecl {
        name: public_name,
        params: params.clone(),
        return_type: return_type.clone(),
        body: wrapped_body,
        modifiers: modifiers.clone(),
        handles: handles.clone(),
        is_async: *is_async,
        is_generator: *is_generator,
        is_sub: *is_sub })
}

fn stmts_to_class_members(class_name: &str, stmts: Vec<Statement>) -> Vec<ClassMember> {
    let mut members: Vec<ClassMember> = Vec::new();
    // Track Property member index by name so @x.setter can find the getter.
    let mut property_indices: std::collections::HashMap<String, usize> =
        std::collections::HashMap::new();

    for stmt in stmts {
        let stmt = if let Some(recovered) = block_desugared_function(&stmt) {
            Statement::new(recovered)
        } else {
            stmt
        };
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                body,
                modifiers,
                is_async,
                ..
            } => {
                if name == "__init__" {
                    // Constructor — keep `self` param; compiler strips it
                    // via `NormalClass.explicit_self_param` (set by
                    // normalize_class).
                    members.push(ClassMember::Constructor {
                        // Python has one constructor, `__init__` — unnamed.
                        name: None,
                        params: params.clone(),
                        body: body.clone(),
                        base_args: None,
                        initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
                        visibility: Visibility::Public });
                    continue;
                }

                // Check for @property decorator → build Property getter
                let has_property = modifiers
                    .decorators
                    .iter()
                    .any(|d| matches!(&d.kind, ExprKind::Ident(n) if n == "property"));
                if has_property {
                    let idx = members.len();
                    members.push(ClassMember::Property {
                        name: name.clone(),
                        type_hint: None,
                        getter: Some(body.clone()),
                        setter: None,
                        is_auto: false,
                        modifiers: Modifiers::default() });
                    property_indices.insert(name.clone(), idx);
                    continue;
                }

                // Check for @x.setter or @x.deleter → add to existing Property
                let setter_target = modifiers.decorators.iter().find_map(|d| {
                    if let ExprKind::Member { object, field, .. } = &d.kind {
                        if field == "setter" {
                            if let ExprKind::Ident(prop_name) = &object.kind {
                                return Some((prop_name.clone(), "setter"));
                            }
                        }
                    }
                    None
                });
                if let Some((prop_name, "setter")) = setter_target {
                    if let Some(&prop_idx) = property_indices.get(&prop_name) {
                        if let ClassMember::Property { setter, .. } = &mut members[prop_idx] {
                            // Second param (after self) is the value param
                            let value_param = params.iter().nth(1).cloned().unwrap_or(Param {
                                name: "value".to_string(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false });
                            *setter = Some(vybe_ast::PropertySetter {
                                param: value_param,
                                body: body.clone() });
                        }
                    }
                    continue;
                }

                // Method — keep `self`/`cls` param; compiler strips via
                // `NormalClass.explicit_self_param`.
                let has_staticmethod = modifiers
                    .decorators
                    .iter()
                    .any(|d| matches!(&d.kind, ExprKind::Ident(n) if n == "staticmethod"));
                let has_classmethod = modifiers
                    .decorators
                    .iter()
                    .any(|d| matches!(&d.kind, ExprKind::Ident(n) if n == "classmethod"));
                let general_decorators: Vec<Expression> = modifiers
                    .decorators
                    .iter()
                    .filter(|d| !is_special_decorator(d))
                    .cloned()
                    .collect();
                let final_body = if general_decorators.is_empty() {
                    body.clone()
                } else {
                    decorated_method_body(
                        name,
                        params,
                        return_type,
                        body,
                        general_decorators,
                    )
                };
                // For @staticmethod, prepend a dummy "self" so that
                // explicit_self_param's skip(1) removes the dummy, keeping
                // the real params intact. Without this, skip(1) would drop
                // the first real param (e.g. `a` in `def add(a, b)`).
                let final_params = if has_staticmethod {
                    let dummy = Param {
                        name: "self".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false };
                    let mut p = vec![dummy];
                    p.extend_from_slice(params);
                    p
                } else if has_classmethod {
                    let dummy = Param {
                        name: "self".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false };
                    let mut p = vec![dummy];
                    if let Some(cls) = params.first() {
                        let mut cls = cls.clone();
                        cls.default = Some(Expression::new(ExprKind::Ident(class_name.to_string())));
                        cls.is_optional = true;
                        p.push(cls);
                        p.extend_from_slice(&params[1..]);
                    }
                    p
                } else {
                    params.clone()
                };
                let is_static =
                    has_staticmethod || has_classmethod || final_params.first().map_or(true, |p| p.name != "self");
                let mut mods = modifiers.clone();
                mods.decorators.retain(is_special_decorator);
                mods.is_static = is_static;
                members.push(ClassMember::Method(Box::new(Statement::new(
                    StmtKind::FunctionDecl {
                        name: name.clone(),
                        params: final_params,
                        return_type: None,
                        body: final_body,
                        modifiers: mods,
                        handles: Vec::new(),
                        is_async: *is_async,
                        is_generator: false,
                        is_sub: false },
                ))));
            }
            // Annotated class-level declaration (`x: int = 0`, or bare
            // `x: int`). The type hint is what marks it a dataclass field, so
            // it is carried onto the member rather than dropped.
            //
            // Only the form WITH a value is a member at all. A bare
            // `x: int` binds nothing in CPython — `P.x` raises
            // AttributeError — it just records `__annotations__` and lets
            // `__init__` create the attribute. Emitting it as a static class
            // attribute made that empty attribute shadow the instance write,
            // so `p = P(1); p.x` read `None` while `P(1).x` read `1`. The
            // dataclass field list is derived from the body statements
            // (`dataclass_fields`), not from these members, so dropping the
            // value-less form costs nothing there.
            StmtKind::VarDecl { declarations, .. } => {
                for d in declarations {
                    let BindingPattern::Ident(field_name) = &d.pattern else {
                        continue;
                    };
                    if d.init.is_none() {
                        continue;
                    }
                    let mut mods = Modifiers::default();
                    mods.is_static = true;
                    members.push(ClassMember::Field {
                        name: field_name.clone(),
                        type_hint: d.type_hint.clone(),
                        init: d.init.clone(),
                        modifiers: mods,
                        with_events: false,
                        array_bounds: None });
                }
            }
            StmtKind::Assign { targets, value , ..} => {
                // Class-level assignment → static Field (Python class variables)
                for target in targets {
                    if let ExprKind::Ident(field_name) = &target.kind {
                        let mut mods = Modifiers::default();
                        mods.is_static = true; // Python class-level vars are class attributes
                        members.push(ClassMember::Field {
                            name: field_name.clone(),
                            type_hint: None,
                            init: Some(value.clone()),
                            modifiers: mods,
                            with_events: false,
                            array_bounds: None });
                    }
                }
            }
            StmtKind::ClassDecl { .. } => {
                members.push(ClassMember::NestedType(Box::new(stmt)));
            }
            StmtKind::Empty => {} // pass
            _ => {
                // Nested class or other — wrap as method
                members.push(ClassMember::Method(Box::new(stmt)));
            }
        }
    }
    members
}

// ── Decorated ───────────────────────────────────────────────────────────────

/// A decorator with dedicated compile-time handling (class-member kind,
/// dataclass, generator, …) rather than Python's runtime `f = deco(f)` wrapping.
/// These stay on the declaration's modifiers for the specialized paths.
fn is_special_decorator(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => matches!(
            name.as_str(),
            "property"
                | "staticmethod"
                | "classmethod"
                | "abstractmethod"
                | "dataclass"
                | "generator"
                | "final"
                | "override"
        ),
        ExprKind::Member { field, .. } => {
            matches!(field.as_str(), "setter" | "getter" | "deleter")
        }
        _ => false }
}

fn function_docstring(body: &[Statement]) -> Option<String> {
    let first = body.first()?;
    if let StmtKind::Expr(expr) = &first.kind
        && let ExprKind::Lit(Literal::Str(doc)) = &expr.kind
    {
        return Some(doc.to_string());
    }
    None
}

fn assign_member(object: Expression, field: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: field.to_string(),
            null_safe: false })],
        value, by_ref: false })
}

fn function_doc_expr(body: &[Statement]) -> Expression {
    function_docstring(body)
        .map(|doc| Expression::string(&doc))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))
}

fn py_annotation_expr(type_hint: &str) -> Expression {
    match type_hint.trim() {
        "int" | "str" | "bool" | "float" | "list" | "dict" | "tuple" | "set" => call_ident(
            "__py_type_obj",
            vec![Expression::string(type_hint.trim())],
        ),
        other => Expression::string(other) }
}

fn function_annotations_expr(params: &[Param], return_type: Option<&String>) -> Option<Expression> {
    let mut props = Vec::new();
    for p in params {
        if let Some(hint) = &p.type_hint {
            props.push(ObjectProperty::KeyValue {
                key: Expression::string(&p.name),
                value: py_annotation_expr(hint) });
        }
    }
    if let Some(hint) = return_type {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string("return"),
            value: py_annotation_expr(hint) });
    }
    if props.is_empty() {
        None
    } else {
        Some(Expression::new(ExprKind::Object(props)))
    }
}

fn assign_function_metadata(
    out: &mut Vec<Statement>,
    fn_name: &str,
    params: &[Param],
    return_type: Option<&String>,
    body: &[Statement],
) {
    out.push(assign_member(
        Expression::ident(fn_name),
        "__name__",
        Expression::string(fn_name),
    ));
    out.push(assign_member(
        Expression::ident(fn_name),
        "__doc__",
        function_doc_expr(body),
    ));
    if let Some(annotations) = function_annotations_expr(params, return_type) {
        out.push(assign_member(
            Expression::ident(fn_name),
            "__annotations__",
            annotations,
        ));
    }
}

fn call_decorator_stack(decorators: Vec<Expression>, base: Expression) -> Expression {
    let mut acc = base;
    for d in decorators.into_iter().rev() {
        acc = Expression::new(call_or_new(
            d,
            vec![Argument {
                value: acc,
                name: None,
                by_ref: false,
                spread: false }],
        ));
    }
    acc
}

fn decorator_root_ident(expr: &Expression) -> Option<&str> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.as_str()),
        ExprKind::Call { callee, .. } => decorator_root_ident(callee),
        ExprKind::Member { object, .. } => decorator_root_ident(object),
        _ => None }
}

fn decorator_stack_contains_class(decorators: &[Expression]) -> bool {
    decorators
        .iter()
        .filter_map(decorator_root_ident)
        .any(is_defined_class)
}

/// Desugar Python function decorators to runtime application:
/// `@a @b def f(...)` → `f = a(b(<function f>))`. Fires only when every
/// decorator is a general (user) decorator; if any is special the declaration
/// is returned unchanged so the specialized compile paths still see it.
fn desugar_function_decorators(decl: StmtKind, decorators: Vec<Expression>) -> StmtKind {
    if decorators.is_empty() || decorators.iter().any(is_special_decorator) {
        return decl;
    }
    let (fn_name, params, return_type, body) = if let StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        ..
    } = &decl {
        (
            name.clone(),
            params.clone(),
            return_type.clone(),
            body.clone(),
        )
    } else {
        return decl;
    };
    // Strip the now-runtime-applied decorators off the inner declaration so the
    // metadata pass doesn't ALSO treat them as (inert) annotations.
    let use_private_original = decorator_stack_contains_class(&decorators);
    let original_name = if use_private_original {
        format!("__py_orig_{fn_name}")
    } else {
        fn_name.clone()
    };
    let mut inner = decl;
    if let StmtKind::FunctionDecl {
        name, modifiers, ..
    } = &mut inner
    {
        if use_private_original {
            *name = original_name.clone();
        }
        modifiers.decorators = Vec::new();
    }
    let mut stmts = vec![Statement::new(inner)];
    assign_function_metadata(&mut stmts, &original_name, &params, return_type.as_ref(), &body);
    if use_private_original {
        stmts.push(assign_member(
            Expression::ident(&original_name),
            "__name__",
            Expression::string(&fn_name),
        ));
    }
    // Fold innermost-first (reversed) so `@a @b def f` becomes `a(b(f))`.
    let acc = call_decorator_stack(decorators, Expression::ident(&original_name));
    stmts.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&fn_name)],
        value: acc, by_ref: false }));
    StmtKind::Block(stmts)
}

fn class_decl_name(decl: &StmtKind) -> Option<String> {
    match decl {
        StmtKind::ClassDecl { name, .. } => Some(name.clone()),
        _ => None }
}

fn desugar_class_decorators(decl: StmtKind, decorators: Vec<Expression>) -> StmtKind {
    let general: Vec<Expression> = decorators
        .into_iter()
        .filter(|d| !is_dataclass_decorator(d))
        .collect();
    if general.is_empty() {
        return decl;
    }
    let Some(name) = class_decl_name(&decl) else {
        return decl;
    };
    let value = call_decorator_stack(general, Expression::ident(&name));
    StmtKind::Block(vec![
        Statement::new(decl),
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&name)],
            value, by_ref: false }),
    ])
}

fn walk_decorated(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut decorators = Vec::new();
    let mut inner_pairs: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Collect decorators
    let i = 0;
    while i < inner_pairs.len() {
        if inner_pairs[i].as_rule() == Rule::decorator {
            let dec_pair = inner_pairs.remove(i);
            for dp in dec_pair.into_inner() {
                if dp.as_rule() != Rule::NEWLINE {
                    decorators.push(walk_expression(dp)?);
                }
            }
        } else {
            break;
        }
    }

    // Remaining should be the def/class
    if let Some(item) = inner_pairs.into_iter().next() {
        match item.as_rule() {
            Rule::function_def => {
                let decl = walk_func_def(item, false, decorators.clone())?;
                Ok(desugar_function_decorators(decl, decorators))
            }
            Rule::class_def => {
                let decl = walk_class_def(item, decorators.clone())?;
                Ok(desugar_class_decorators(decl, decorators))
            }
            Rule::async_stmt => {
                // async def with decorators
                for p in item.into_inner() {
                    match p.as_rule() {
                        Rule::function_def => {
                            let decl = walk_func_def(p, true, decorators.clone())?;
                            return Ok(desugar_function_decorators(decl, decorators));
                        }
                        Rule::for_stmt => return walk_for(p, true),
                        Rule::with_stmt => return walk_with(p, true),
                        _ => {}
                    }
                }
                Err("Expected def/for/with after async".into())
            }
            other => Err(format!(
                "Expected def/class after decorator, got {:?}",
                other
            )) }
    } else {
        Err("Empty decorated statement".into())
    }
}

// ── Async ───────────────────────────────────────────────────────────────────

fn walk_async_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::function_def => return walk_func_def(p, true, Vec::new()),
            Rule::for_stmt => return walk_for(p, true),
            Rule::with_stmt => return walk_with(p, true),
            Rule::async_kw => {}
            _ => {}
        }
    }
    Err("Expected def/for/with after async".into())
}

// ── If ──────────────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let then_body = walk_block(next_rule_any(&mut inner, &[Rule::block])?)?;

    let mut elifs = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::elif_clause => {
                let mut ei = p.into_inner();
                let econd = walk_expression(next_meaningful(&mut ei)?)?;
                let ebody = walk_block(next_rule_any(&mut ei, &[Rule::block])?)?;
                elifs.push((econd, ebody));
            }
            Rule::else_clause => {
                let mut ei = p.into_inner();
                else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body })
}

// ── While ───────────────────────────────────────────────────────────────────

/// Per-parse counter that keeps desugared `while…else` break-flags unique so
/// nested loops don't share (and clobber) one flag.
static WHILE_ELSE_COUNTER: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

/// Rewrite loop-level `break` statements in `stmts` to set `flag = True` before
/// breaking, so a desugared `while…else` can distinguish a break-exit from a
/// normal exit. Recurses through non-loop containers (if/try/with/block) but NOT
/// into nested loops or function/class bodies — their `break`s target
/// themselves, not this loop.
fn mark_loop_break_sets_flag(stmts: Vec<Statement>, flag: &str) -> Vec<Statement> {
    let mut out = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        match stmt.kind {
            StmtKind::Break(_) => {
                out.push(Statement::new(StmtKind::Assign {
                    targets: vec![Expression::ident(flag)],
                    value: Expression::bool(true), by_ref: false }));
                out.push(stmt);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body } => {
                out.push(Statement::new(StmtKind::If {
                    cond,
                    then_body: mark_loop_break_sets_flag(then_body, flag),
                    elifs: elifs
                        .into_iter()
                        .map(|(c, b)| (c, mark_loop_break_sets_flag(b, flag)))
                        .collect(),
                    else_body: else_body.map(|b| mark_loop_break_sets_flag(b, flag)) }));
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally } => {
                out.push(Statement::new(StmtKind::Try {
                    body: mark_loop_break_sets_flag(body, flag),
                    catches: catches
                        .into_iter()
                        .map(|mut c| {
                            c.body = mark_loop_break_sets_flag(c.body, flag);
                            c
                        })
                        .collect(),
                    else_body: else_body.map(|b| mark_loop_break_sets_flag(b, flag)),
                    finally: finally.map(|b| mark_loop_break_sets_flag(b, flag)) }));
            }
            StmtKind::With {
                items,
                body,
                is_async } => {
                out.push(Statement::new(StmtKind::With {
                    items,
                    body: mark_loop_break_sets_flag(body, flag),
                    is_async }));
            }
            StmtKind::Block(b) => {
                out.push(Statement::new(StmtKind::Block(mark_loop_break_sets_flag(
                    b, flag,
                ))));
            }
            _ => out.push(stmt) }
    }
    out
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let body = walk_block(next_rule_any(&mut inner, &[Rule::block])?)?;
    let mut else_body = None;
    for p in inner {
        if p.as_rule() == Rule::else_clause {
            let mut ei = p.into_inner();
            else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
        }
    }

    let Some(else_stmts) = else_body else {
        return Ok(StmtKind::While {
            cond,
            body,
            else_body: None });
    };

    // Python `while C: BODY else: ELSE` runs ELSE only on a NORMAL exit
    // (condition false), never on `break`. The shared While emitter runs
    // else_body unconditionally, so normalize into common-AST primitives that
    // route through the common loop emitter (loops.rs), the same plain-`while`
    // path every language uses:
    //   __while_else_N = False
    //   while C: BODY'                 (loop-level break → __while_else_N = True; break)
    //   if not __while_else_N: ELSE
    let n = WHILE_ELSE_COUNTER.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let flag = format!("__while_else_{n}");
    let body = mark_loop_break_sets_flag(body, &flag);
    let flag_init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&flag)],
        value: Expression::bool(false), by_ref: false });
    let while_stmt = Statement::new(StmtKind::While {
        cond,
        body,
        else_body: None });
    let else_guard = Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(Expression::ident(&flag)) }),
        then_body: else_stmts,
        elifs: Vec::new(),
        else_body: None });
    Ok(StmtKind::Block(vec![flag_init, while_stmt, else_guard]))
}

// ── For ─────────────────────────────────────────────────────────────────────

fn walk_for(pair: Pair<Rule>, is_async: bool) -> Result<StmtKind, String> {
    let inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Find target_list, expression_list, block, else_clause
    let mut var_names: Vec<String> = Vec::new();
    let mut iter_expr = None;
    let mut body = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::target_list => {
                let text = p.as_str().trim().to_string();
                var_names = text
                    .split(',')
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty())
                    .collect();
            }
            Rule::expression_list => {
                if iter_expr.is_none() {
                    iter_expr = Some(walk_expr_list(p)?);
                }
            }
            Rule::block => body = walk_block(p)?,
            Rule::else_clause => {
                let mut ei = p.into_inner();
                else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
            }
            Rule::in_kw | Rule::async_kw => {}
            _ => {}
        }
    }

    // If multiple targets (tuple unpacking: `for i, v in enumerate(...)`),
    // use a temp var and prepend destructuring assignments to the body.
    let var = if var_names.len() > 1 {
        let tmp = "__forin_element".to_string();
        let mut destructure_stmts: Vec<Statement> = Vec::new();
        for (i, name) in var_names.iter().enumerate() {
            // name = __forin_element[i]
            destructure_stmts.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(name.clone()))],
                value: Expression::new(ExprKind::Index {
                    object: Box::new(Expression::new(ExprKind::Ident(tmp.clone()))),
                    index: Box::new(Expression::new(ExprKind::Lit(Literal::Int(i as i64)))),
                    null_safe: false }), by_ref: false }));
        }
        destructure_stmts.extend(body);
        body = destructure_stmts;
        tmp
    } else {
        var_names.into_iter().next().unwrap_or_default()
    };

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter: iter_expr.unwrap_or(Expression::new(ExprKind::Lit(Literal::Null))),
        body,
        of: true,
        else_body,
        is_async })
}

// ── Try ─────────────────────────────────────────────────────────────────────

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut else_body = None;
    let mut finally = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block => {
                if body.is_empty() {
                    body = walk_block(p)?;
                }
            }
            Rule::except_clause => {
                let mut types = Vec::new();
                let mut var_name = None;
                let mut catch_body = Vec::new();
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::block => catch_body = walk_block(cp)?,
                        Rule::identifier => var_name = Some(cp.as_str().to_string()),
                        Rule::as_kw => {}
                        _ => {
                            // Exception type expression. Python allows tuple
                            // catches (`except (A, B):`); the common catch
                            // node wants one type name per entry.
                            types.extend(py_except_type_names(cp.as_str()));
                        }
                    }
                }
                if catch_body.iter().any(py_stmt_is_bare_raise) {
                    let rethrow_var = var_name
                        .clone()
                        .unwrap_or_else(|| "__py_current_exception".to_string());
                    py_rewrite_bare_raises(&mut catch_body, &rethrow_var);
                    var_name = Some(rethrow_var);
                }
                let current_var = var_name
                    .clone()
                    .unwrap_or_else(|| "__py_current_exception".to_string());
                py_expand_except_exc_info_assigns(&mut catch_body, &current_var);
                py_rewrite_except_exc_info(&mut catch_body, &current_var);
                py_stamp_except_raise_contexts(&mut catch_body, &current_var);
                if var_name.is_none() {
                    var_name = Some(current_var);
                }
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None });
            }
            Rule::else_clause => {
                let mut ei = p.into_inner();
                else_body = Some(walk_block(next_rule_any(&mut ei, &[Rule::block])?)?);
            }
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block {
                        finally = Some(walk_block(fp)?);
                    }
                }
            }
            _ => {}
        }
    }

    let has_later_base_exception = |idx: usize, catches: &[CatchClause]| {
        catches
            .iter()
            .skip(idx + 1)
            .any(|catch| catch.types.iter().any(|ty| ty == "BaseException"))
    };
    for idx in 0..catches.len() {
        if has_later_base_exception(idx, &catches)
            && catches[idx].types.iter().any(|ty| ty == "Exception")
        {
            catches[idx].types.retain(|ty| ty != "Exception");
        }
    }

    Ok(StmtKind::Try {
        body,
        catches,
        else_body,
        finally })
}

fn py_except_type_names(raw: &str) -> Vec<String> {
    let text = raw.trim();
    let inner = text
        .strip_prefix('(')
        .and_then(|s| s.strip_suffix(')'))
        .unwrap_or(text);
    let mut out = Vec::new();
    for name in inner
        .split(',')
        .map(str::trim)
        .filter(|s| !s.is_empty())
    {
        let normalized = name
            .rsplit('.')
            .next()
            .filter(|leaf| py_builtin_exception_bases(leaf).is_some())
            .unwrap_or(name);
        out.push(normalized.to_string());
        if normalized == "GetoptError" && !out.iter().any(|existing| existing == "Exception") {
            out.push("Exception".to_string());
        }
        if py_builtin_exception_bases(normalized).is_some() {
            for candidate in py_builtin_exception_names() {
                if *candidate != normalized
                    && py_builtin_subclass(candidate, normalized) == Some(true)
                    && !out.iter().any(|existing| existing == candidate)
                {
                    out.push((*candidate).to_string());
                }
            }
        }
    }
    out
}

fn py_stmt_is_bare_raise(stmt: &Statement) -> bool {
    matches!(stmt.kind, StmtKind::Throw { expr: None, .. })
}

fn py_rewrite_bare_raises(body: &mut [Statement], var_name: &str) {
    for stmt in body {
        if let StmtKind::Throw { expr, .. } = &mut stmt.kind
            && expr.is_none()
        {
            *expr = Some(Expression::ident(var_name));
        }
    }
}

fn py_is_ident(expr: &Expression, name: &str) -> bool {
    matches!(&expr.kind, ExprKind::Ident(n) if n == name)
}

fn py_exception_chain_stmts(
    exc: Expression,
    cause: Expression,
    context: Expression,
    suppress: bool,
) -> Vec<Statement> {
    let tmp = "__py_pending_exception";
    vec![
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(tmp)],
            value: exc, by_ref: false }),
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(tmp)),
                field: "__cause__".into(),
                null_safe: false })],
            value: cause, by_ref: false }),
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(tmp)),
                field: "__context__".into(),
                null_safe: false })],
            value: context, by_ref: false }),
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(tmp)),
                field: "__suppress_context__".into(),
                null_safe: false })],
            value: Expression::bool(suppress), by_ref: false }),
        Statement::new(StmtKind::Throw {
            expr: Some(Expression::ident(tmp)),
            cause: None }),
    ]
}

fn py_stamp_except_raise_contexts(body: &mut Vec<Statement>, current_var: &str) {
    let mut rewritten = Vec::with_capacity(body.len());
    for mut stmt in body.drain(..) {
        let replacement = match &mut stmt.kind {
            StmtKind::Expr(Expression {
                kind: ExprKind::Call { callee, args, .. },
                ..
            }) => {
                if let ExprKind::Ident(name) = &callee.kind
                    && let Some(exc_name) = name.strip_prefix("__py_raise_")
                {
                    let exc = call_ident(
                        &format!("__py_exc_{exc_name}"),
                        args.iter().map(|a| a.value.clone()).collect(),
                    );
                    Some(py_exception_chain_stmts(
                        exc,
                        Expression::new(ExprKind::Lit(Literal::Null)),
                        Expression::ident(current_var),
                        false,
                    ))
                } else {
                    None
                }
            }
            StmtKind::Throw { expr, cause } => {
                if let Some(exc) = expr.take() {
                    if py_is_ident(&exc, current_var) && cause.is_none() {
                        *expr = Some(exc);
                        None
                    } else {
                        let (cause_expr, suppress) = if let Some(cause) = cause.take() {
                            (cause, true)
                        } else {
                            (Expression::new(ExprKind::Lit(Literal::Null)), false)
                        };
                        Some(py_exception_chain_stmts(
                            exc,
                            cause_expr,
                            Expression::ident(current_var),
                            suppress,
                        ))
                    }
                } else {
                    None
                }
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                py_stamp_except_raise_contexts(then_body, current_var);
                for (_, body) in elifs {
                    py_stamp_except_raise_contexts(body, current_var);
                }
                if let Some(body) = else_body {
                    py_stamp_except_raise_contexts(body, current_var);
                }
                None
            }
            StmtKind::Block(stmts) => {
                py_stamp_except_raise_contexts(stmts, current_var);
                None
            }
            StmtKind::While { body, else_body, .. } | StmtKind::ForIn { body, else_body, .. } => {
                py_stamp_except_raise_contexts(body, current_var);
                if let Some(body) = else_body {
                    py_stamp_except_raise_contexts(body, current_var);
                }
                None
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally } => {
                py_stamp_except_raise_contexts(body, current_var);
                for catch in catches {
                    py_stamp_except_raise_contexts(&mut catch.body, current_var);
                }
                if let Some(body) = else_body {
                    py_stamp_except_raise_contexts(body, current_var);
                }
                if let Some(body) = finally {
                    py_stamp_except_raise_contexts(body, current_var);
                }
                None
            }
            _ => None };
        if let Some(stmts) = replacement {
            rewritten.extend(stmts);
        } else {
            rewritten.push(stmt);
        }
    }
    *body = rewritten;
}

fn py_rewrite_except_exc_info(body: &mut [Statement], current_var: &str) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                py_rewrite_except_exc_info_expr(expr, current_var)
            }
            StmtKind::Assign { targets, value , ..} => {
                for target in targets {
                    py_rewrite_except_exc_info_expr(target, current_var);
                }
                py_rewrite_except_exc_info_expr(value, current_var);
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        py_rewrite_except_exc_info_expr(init, current_var);
                    }
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body } => {
                py_rewrite_except_exc_info_expr(cond, current_var);
                py_rewrite_except_exc_info(then_body, current_var);
                for (cond, body) in elifs {
                    py_rewrite_except_exc_info_expr(cond, current_var);
                    py_rewrite_except_exc_info(body, current_var);
                }
                if let Some(body) = else_body {
                    py_rewrite_except_exc_info(body, current_var);
                }
            }
            StmtKind::While { cond, body, else_body } => {
                py_rewrite_except_exc_info_expr(cond, current_var);
                py_rewrite_except_exc_info(body, current_var);
                if let Some(body) = else_body {
                    py_rewrite_except_exc_info(body, current_var);
                }
            }
            StmtKind::ForIn {
                iter,
                body,
                else_body,
                ..
            } => {
                py_rewrite_except_exc_info_expr(iter, current_var);
                py_rewrite_except_exc_info(body, current_var);
                if let Some(body) = else_body {
                    py_rewrite_except_exc_info(body, current_var);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally } => {
                py_rewrite_except_exc_info(body, current_var);
                for catch in catches {
                    let nested = catch.var_name.as_deref().unwrap_or(current_var);
                    py_rewrite_except_exc_info(&mut catch.body, nested);
                }
                if let Some(body) = else_body {
                    py_rewrite_except_exc_info(body, current_var);
                }
                if let Some(body) = finally {
                    py_rewrite_except_exc_info(body, current_var);
                }
            }
            StmtKind::Throw { expr, cause } => {
                if let Some(expr) = expr {
                    py_rewrite_except_exc_info_expr(expr, current_var);
                }
                if let Some(cause) = cause {
                    py_rewrite_except_exc_info_expr(cause, current_var);
                }
            }
            _ => {}
        }
    }
}

fn py_expand_except_exc_info_assigns(body: &mut Vec<Statement>, current_var: &str) {
    let mut rewritten = Vec::with_capacity(body.len());
    for mut stmt in body.drain(..) {
        if let StmtKind::Assign { targets, value , ..} = &stmt.kind
            && targets.len() == 1
            && py_is_sys_exc_info_call(value)
            && let Some(names) = py_destructure_idents(&targets[0])
            && names.len() >= 3
        {
            rewritten.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&names[0])],
                value: Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                    key: Expression::string("__name__"),
                    value: call_ident("__py_type_name", vec![Expression::ident(current_var)]) }])), by_ref: false }));
            rewritten.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&names[1])],
                value: Expression::ident(current_var), by_ref: false }));
            rewritten.push(Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&names[2])],
                value: Expression::new(ExprKind::Lit(Literal::Null)), by_ref: false }));
            continue;
        }
        match &mut stmt.kind {
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                py_expand_except_exc_info_assigns(then_body, current_var);
                for (_, body) in elifs {
                    py_expand_except_exc_info_assigns(body, current_var);
                }
                if let Some(body) = else_body {
                    py_expand_except_exc_info_assigns(body, current_var);
                }
            }
            StmtKind::Block(stmts) => py_expand_except_exc_info_assigns(stmts, current_var),
            StmtKind::While { body, else_body, .. } | StmtKind::ForIn { body, else_body, .. } => {
                py_expand_except_exc_info_assigns(body, current_var);
                if let Some(body) = else_body {
                    py_expand_except_exc_info_assigns(body, current_var);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally } => {
                py_expand_except_exc_info_assigns(body, current_var);
                for catch in catches {
                    let nested = catch.var_name.as_deref().unwrap_or(current_var);
                    py_expand_except_exc_info_assigns(&mut catch.body, nested);
                }
                if let Some(body) = else_body {
                    py_expand_except_exc_info_assigns(body, current_var);
                }
                if let Some(body) = finally {
                    py_expand_except_exc_info_assigns(body, current_var);
                }
            }
            _ => {}
        }
        rewritten.push(stmt);
    }
    *body = rewritten;
}

fn py_is_sys_exc_info_call(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, args, .. }
            if args.is_empty()
                && matches!(
                    &callee.kind,
                    ExprKind::Member { object, field, .. }
                        if field == "exc_info"
                            && matches!(&object.kind, ExprKind::Ident(name) if name == "sys")
                )
    )
}

fn py_destructure_idents(expr: &Expression) -> Option<Vec<String>> {
    let ExprKind::Destructure(DestructurePattern::Array(elems)) = &expr.kind else {
        return None;
    };
    let mut names = Vec::new();
    for elem in elems {
        let ArrayPatternElem::Pattern(BindingPattern::Ident(name), None) = elem else {
            return None;
        };
        names.push(name.clone());
    }
    Some(names)
}

fn py_rewrite_except_exc_info_expr(expr: &mut Expression, current_var: &str) {
    match &mut expr.kind {
        ExprKind::Index { object, index, .. }
            if py_is_sys_exc_info_call(object)
                && matches!(&index.kind, ExprKind::Lit(Literal::Int(0 | 1 | 2))) =>
        {
            let ExprKind::Lit(Literal::Int(which)) = index.kind else {
                return;
            };
            expr.kind = match which {
                0 => ExprKind::Object(vec![ObjectProperty::KeyValue {
                    key: Expression::string("__name__"),
                    value: call_ident("__py_type_name", vec![Expression::ident(current_var)]) }]),
                1 => ExprKind::Ident(current_var.to_string()),
                _ => ExprKind::Lit(Literal::Null) };
        }
        ExprKind::Call { callee, args, .. } => {
            if args.is_empty()
                && let ExprKind::Member { object, field, .. } = &callee.kind
                && field == "exc_info"
                && matches!(&object.kind, ExprKind::Ident(name) if name == "sys")
            {
                expr.kind = ExprKind::Array(vec![
                    ArrayElement {
                        key: None,
                        value: Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                            key: Expression::string("__name__"),
                            value: call_ident(
                                "__py_type_name",
                                vec![Expression::ident(current_var)],
                            ) }])),
                        spread: false,
                        by_ref: false },
                    ArrayElement {
                        key: None,
                        value: Expression::ident(current_var),
                        spread: false,
                        by_ref: false },
                    ArrayElement {
                        key: None,
                        value: Expression::new(ExprKind::Lit(Literal::Null)),
                        spread: false,
                        by_ref: false },
                ]);
                return;
            }
            py_rewrite_except_exc_info_expr(callee, current_var);
            for arg in args {
                py_rewrite_except_exc_info_expr(&mut arg.value, current_var);
            }
        }
        ExprKind::Member { object, .. } => py_rewrite_except_exc_info_expr(object, current_var),
        ExprKind::Index { object, index, .. } => {
            py_rewrite_except_exc_info_expr(object, current_var);
            py_rewrite_except_exc_info_expr(index, current_var);
        }
        ExprKind::Binary { left, right, .. } => {
            py_rewrite_except_exc_info_expr(left, current_var);
            py_rewrite_except_exc_info_expr(right, current_var);
        }
        ExprKind::Unary { expr, .. } => py_rewrite_except_exc_info_expr(expr, current_var),
        ExprKind::Ternary { cond, then, else_ } => {
            py_rewrite_except_exc_info_expr(cond, current_var);
            py_rewrite_except_exc_info_expr(then, current_var);
            py_rewrite_except_exc_info_expr(else_, current_var);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    py_rewrite_except_exc_info_expr(key, current_var);
                }
                py_rewrite_except_exc_info_expr(&mut item.value, current_var);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                py_rewrite_except_exc_info_expr(item, current_var);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        py_rewrite_except_exc_info_expr(key, current_var);
                        py_rewrite_except_exc_info_expr(value, current_var);
                    }
                    ObjectProperty::Spread(value) => {
                        py_rewrite_except_exc_info_expr(value, current_var);
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

// ── With ────────────────────────────────────────────────────────────────────

fn walk_with(pair: Pair<Rule>, is_async: bool) -> Result<StmtKind, String> {
    let mut items = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::with_item => {
                let mut expr = None;
                let mut var = None;
                for wp in p.into_inner() {
                    match wp.as_rule() {
                        Rule::as_kw => {}
                        Rule::target | Rule::target_list => {
                            var = Some(wp.as_str().trim().to_string())
                        }
                        _ => {
                            if expr.is_none() {
                                expr = Some(walk_expression(wp)?);
                            }
                        }
                    }
                }
                items.push(WithItem {
                    expr: expr.unwrap_or(Expression::new(ExprKind::Lit(Literal::Null))),
                    var });
            }
            Rule::block => body = walk_block(p)?,
            _ => {}
        }
    }

    // Desugar `with` to the PEP-343 try/except/finally so the standard Try
    // compilation drives the context-manager protocol — reusing errors.rs's
    // try_table AND the finally machinery that already runs on break/continue/
    // return/exception paths (a hand-rolled try_table hangs on break/nested).
    let _ = is_async;
    if items.is_empty() {
        return Ok(StmtKind::Block(body));
    }
    Ok(StmtKind::Block(build_with_desugar(&items, body)))
}

use std::sync::atomic::{AtomicUsize, Ordering};
static WITH_COUNTER: AtomicUsize = AtomicUsize::new(0);

fn with_stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}
fn with_member(recv: &str, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(recv)),
        field: field.to_string(),
        null_safe: false })
}
fn with_arg(value: Expression) -> Argument {
    Argument {
        value,
        name: None,
        by_ref: false,
        spread: false }
}
fn with_call(callee: Expression, args: Vec<Argument>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false })
}
fn with_not(e: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(e) })
}

/// `with A as a, B as b: body` → nested PEP-343 blocks:
/// ```text
/// __mgr = A; a = __mgr.__enter__(); __hit = False
/// try: <inner>
/// except as __e: __hit = True; if not __mgr.__exit__(__e, __e, None): raise
/// finally: if not __hit: __mgr.__exit__(None, None, None)
/// ```
/// `with conn:` for a tracked sqlite3 connection → transaction:
/// `__sql_begin`; on normal exit `__sql_commit`; on exception `__sql_rollback`
/// then re-raise (sqlite3 does NOT suppress the exception).
fn build_sql_with_desugar(conn: &str, body: Vec<Statement>) -> Vec<Statement> {
    let n = WITH_COUNTER.fetch_add(1, Ordering::Relaxed);
    let hit = format!("__sql_hit_{n}");
    let exc = format!("__sql_exc_{n}");

    let sql_call = |name: &str| {
        with_stmt(StmtKind::Expr(with_call(
            Expression::ident(name),
            vec![with_arg(Expression::ident(conn))],
        )))
    };

    let catch = CatchClause {
        types: vec![],
        var_name: Some(exc.clone()),
        stack_var: None,
        body: vec![
            with_stmt(StmtKind::Assign {
                targets: vec![Expression::ident(&hit)],
                value: Expression::bool(true), by_ref: false }),
            sql_call("__sql_rollback"),
            with_stmt(StmtKind::Throw {
                expr: Some(Expression::ident(&exc)),
                cause: None }),
        ],
        when_clause: None };

    let finally = vec![with_stmt(StmtKind::If {
        cond: with_not(Expression::ident(&hit)),
        then_body: vec![sql_call("__sql_commit")],
        elifs: vec![],
        else_body: None })];

    vec![
        sql_call("__sql_begin"),
        with_stmt(StmtKind::Assign {
            targets: vec![Expression::ident(&hit)],
            value: Expression::bool(false), by_ref: false }),
        with_stmt(StmtKind::Try {
            body,
            catches: vec![catch],
            else_body: None,
            finally: Some(finally) }),
    ]
}

/// `with open(...) as f: BODY` → `f = open(...)` + `try: BODY finally: f.close()`.
fn build_file_with_desugar(
    item: &WithItem,
    body: Vec<Statement>,
    closes: bool,
) -> Vec<Statement> {
    let n = WITH_COUNTER.fetch_add(1, Ordering::Relaxed);
    let target = item
        .var
        .clone()
        .unwrap_or_else(|| format!("__with_file_{n}"));
    vec![
        with_stmt(StmtKind::Assign {
            targets: vec![Expression::ident(&target)],
            value: item.expr.clone(), by_ref: false }),
        with_stmt(StmtKind::Try {
            body,
            catches: vec![],
            else_body: None,
            finally: Some(if closes {
                vec![with_stmt(StmtKind::Expr(with_call(
                    Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(&target)),
                        field: "close".into(),
                        null_safe: false }),
                    vec![],
                )))]
            } else {
                vec![]
            }) }),
    ]
}

fn suppress_context_types(expr: &Expression) -> Option<Vec<String>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let is_suppress = match &callee.kind {
        ExprKind::Ident(name) => name == "suppress",
        ExprKind::Member { object, field, .. } => {
            field == "suppress"
                && matches!(&object.kind, ExprKind::Ident(module) if module == "contextlib")
        }
        _ => false };
    if !is_suppress {
        return None;
    }
    let mut types = Vec::new();
    for arg in args {
        if let ExprKind::Ident(name) = &arg.value.kind {
            types.extend(py_except_type_names(name));
        }
    }
    Some(types)
}

fn build_with_desugar(items: &[WithItem], body: Vec<Statement>) -> Vec<Statement> {
    let first = &items[0];
    // sqlite3 Connection used as a context manager → transaction semantics.
    if items.len() == 1 && first.var.is_none() {
        if let ExprKind::Ident(name) = &first.expr.kind {
            if is_sql_var(name) {
                return build_sql_with_desugar(name, body);
            }
        }
    }
    if items.len() == 1
        && first.var.is_none()
        && let Some(types) = suppress_context_types(&first.expr)
    {
        return vec![with_stmt(StmtKind::Try {
            body,
            catches: vec![CatchClause {
                types,
                var_name: None,
                stack_var: None,
                body: Vec::new(),
                when_clause: None }],
            else_body: None,
            finally: None })];
    }
    // `with open(...) as f:` — a file object is a plain adapter-built value, not
    // a class, so it has no `__enter__`/`__exit__` to call. Bind it directly and
    // close in a `finally`, which IS CPython's file context-manager semantics.
    // (Adding `__enter__`/`__exit__` as value_methods instead would shadow every
    // user-defined context manager, since value methods win over user methods.)
    if items.len() == 1
        && let ExprKind::Call { callee, .. } = &first.expr.kind
        && (matches!(&callee.kind, ExprKind::Ident(n) if n == "open")
            // `tempfile.NamedTemporaryFile(...)` yields the same file object.
            || matches!(&callee.kind, ExprKind::Member { object, field, .. }
                if matches!(&object.kind, ExprKind::Ident(m) if m == "tempfile")
                    && (field == "NamedTemporaryFile"
                        || field == "TemporaryFile"
                        || field == "TemporaryDirectory"))
            // `with os.scandir(d) as it:` — an array; nothing to close.
            || matches!(&callee.kind, ExprKind::Member { object, field, .. }
                if matches!(&object.kind, ExprKind::Ident(m) if m == "os")
                    && field == "scandir"))
    {
        // `TemporaryDirectory()` yields a PATH STRING, not a file object — it
        // has no `close()`, so bind it and skip the cleanup call.
        let closes = !matches!(&first.expr.kind, ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Member { field, .. }
                if field == "TemporaryDirectory" || field == "scandir"));
        return build_file_with_desugar(first, body, closes);
    }
    let n = WITH_COUNTER.fetch_add(1, Ordering::Relaxed);
    let mgr = format!("__with_mgr_{n}");
    let hit = format!("__with_hit_{n}");
    let exc = format!("__with_exc_{n}");
    let target = first
        .var
        .clone()
        .unwrap_or_else(|| format!("__with_target_{n}"));

    let inner_body = if items.len() > 1 {
        build_with_desugar(&items[1..], body)
    } else {
        body
    };

    let assign = |name: &str, value: Expression| {
        with_stmt(StmtKind::Assign {
            targets: vec![Expression::ident(name)],
            value, by_ref: false })
    };

    let catch = CatchClause {
        types: vec![],
        var_name: Some(exc.clone()),
        stack_var: None,
        body: vec![
            assign(&hit, Expression::bool(true)),
            with_stmt(StmtKind::If {
                // PEP-343: `__exit__(exc_type, exc_value, traceback)`. Pass the
                // exception TYPE (not the instance) as the first arg so
                // `issubclass(exc_type, …)` works (e.g. contextlib.suppress).
                cond: with_not(with_call(
                    with_member(&mgr, "__exit__"),
                    vec![
                        with_arg(with_call(
                            Expression::ident("type"),
                            vec![with_arg(Expression::ident(&exc))],
                        )),
                        with_arg(Expression::ident(&exc)),
                        with_arg(Expression::null()),
                    ],
                )),
                // `__exit__` returned falsy → re-raise the caught exception. A
                // bare `raise` re-raises `null` here, so raise `exc` explicitly.
                then_body: vec![with_stmt(StmtKind::Throw {
                    expr: Some(Expression::ident(&exc)),
                    cause: None })],
                elifs: vec![],
                else_body: None }),
        ],
        when_clause: None };

    let finally = vec![with_stmt(StmtKind::If {
        cond: with_not(Expression::ident(&hit)),
        then_body: vec![with_stmt(StmtKind::Expr(with_call(
            with_member(&mgr, "__exit__"),
            vec![
                with_arg(Expression::null()),
                with_arg(Expression::null()),
                with_arg(Expression::null()),
            ],
        )))],
        elifs: vec![],
        else_body: None })];

    vec![
        assign(&mgr, first.expr.clone()),
        assign(&target, with_call(with_member(&mgr, "__enter__"), vec![])),
        assign(&hit, Expression::bool(false)),
        with_stmt(StmtKind::Try {
            body: inner_body,
            catches: vec![catch],
            else_body: None,
            finally: Some(finally) }),
    ]
}

// ── Match ───────────────────────────────────────────────────────────────────

fn walk_match(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut subject = None;
    let mut cases = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression_list => {
                if subject.is_none() {
                    subject = Some(walk_expr_list(p)?);
                }
            }
            Rule::case_clause => {
                let mut pattern = Pattern::Wildcard;
                let mut guard = None;
                let mut body = Vec::new();
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::pattern | Rule::or_pattern => pattern = walk_pattern(cp)?,
                        Rule::block => body = walk_block(cp)?,
                        _ => {
                            // Guard expression (after "if")
                            if cp.as_rule() != Rule::if_kw {
                                guard = Some(walk_expression(cp)?);
                            }
                        }
                    }
                }
                cases.push(MatchCase {
                    pattern,
                    guard,
                    body });
            }
            Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => {}
            _ => {}
        }
    }

    Ok(StmtKind::MatchStatement {
        subject: subject.unwrap_or(Expression::new(ExprKind::Lit(Literal::Null))),
        cases })
}

fn walk_pattern(pair: Pair<Rule>) -> Result<Pattern, String> {
    match pair.as_rule() {
        Rule::or_pattern => {
            let pats: Vec<Pair<Rule>> = pair.into_inner().collect();
            if pats.len() == 1 {
                walk_pattern(pats.into_iter().next().unwrap())
            } else {
                let patterns = pats
                    .into_iter()
                    .map(walk_pattern)
                    .collect::<Result<Vec<_>, _>>()?;
                Ok(Pattern::Or(patterns))
            }
        }
        Rule::pattern => {
            let inner = pair.into_inner().next();
            match inner {
                Some(p) => walk_pattern(p),
                None => Ok(Pattern::Wildcard) }
        }
        Rule::single_pattern => {
            let inner = pair.into_inner().next().ok_or("Empty single_pattern")?;
            walk_pattern(inner)
        }
        Rule::group_pattern => {
            let inner = pair.into_inner().next().ok_or("Empty group_pattern")?;
            walk_pattern(inner)
        }
        Rule::as_pattern => {
            // pattern as name
            let mut inner = pair.into_inner();
            let sub_pattern = walk_pattern(inner.next().ok_or("Missing as_pattern sub-pattern")?)?;
            // skip as_kw
            let name = inner
                .filter(|p| p.as_rule() == Rule::identifier)
                .next()
                .map(|p| p.as_str().to_string());
            Ok(Pattern::As {
                pattern: Some(Box::new(sub_pattern)),
                name })
        }
        Rule::wildcard_pattern => Ok(Pattern::Wildcard),
        Rule::capture_pattern => {
            let name = pair.as_str().to_string();
            Ok(Pattern::As {
                pattern: None,
                name: Some(name) })
        }
        Rule::singleton_pattern => {
            let text = pair.as_str().trim();
            let expr = match text {
                "None" => Expression::null(),
                "True" => Expression::bool(true),
                "False" => Expression::bool(false),
                _ => Expression::null() };
            Ok(Pattern::Singleton(expr))
        }
        Rule::literal_pattern => {
            let text = pair.as_str().trim();
            let expr = parse_literal_to_expr(text);
            Ok(Pattern::Value(expr))
        }
        Rule::value_pattern => {
            // Dotted name like module.CONST
            let expr = Expression::new(ExprKind::Ident(pair.as_str().trim().to_string()));
            Ok(Pattern::Value(expr))
        }
        Rule::star_pattern => {
            let name = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::identifier)
                .map(|p| p.as_str().to_string());
            Ok(Pattern::Star(name))
        }
        Rule::sequence_pattern => {
            let pats = pair
                .into_inner()
                .map(walk_pattern)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(Pattern::Sequence(pats))
        }
        Rule::tuple_pattern => {
            let pats = pair
                .into_inner()
                .map(walk_pattern)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(Pattern::Sequence(pats))
        }
        Rule::mapping_pattern => {
            let mut pairs_vec = Vec::new();
            for mp in pair.into_inner() {
                if mp.as_rule() == Rule::mapping_pair {
                    let mut mi = mp.into_inner();
                    let key = walk_expression(mi.next().ok_or("Missing mapping key")?)?;
                    let val = walk_pattern(mi.next().ok_or("Missing mapping pattern")?)?;
                    pairs_vec.push((key, val));
                }
            }
            Ok(Pattern::Mapping(pairs_vec))
        }
        Rule::class_pattern => {
            let mut cls_name = String::new();
            let mut patterns = Vec::new();
            let mut kw_patterns = Vec::new();
            for cp in pair.into_inner() {
                match cp.as_rule() {
                    Rule::identifier => cls_name = cp.as_str().to_string(),
                    Rule::class_pattern_arg => {
                        let mut ai = cp.into_inner();
                        let first = ai.next().ok_or("Empty class_pattern_arg")?;
                        if first.as_rule() == Rule::identifier {
                            // Could be keyword=pattern or just a capture pattern
                            if let Some(second) = ai.next() {
                                // keyword = pattern
                                let name = first.as_str().to_string();
                                let pat = walk_pattern(second)?;
                                kw_patterns.push((name, pat));
                            } else {
                                // Just a pattern (identifier is capture or wildcard)
                                patterns.push(walk_pattern(first)?);
                            }
                        } else {
                            patterns.push(walk_pattern(first)?);
                        }
                    }
                    _ => patterns.push(walk_pattern(cp)?) }
            }
            Ok(Pattern::Class {
                cls: Expression::new(ExprKind::Ident(cls_name)),
                patterns,
                kw_patterns })
        }
        Rule::true_kw => Ok(Pattern::Singleton(Expression::bool(true))),
        Rule::false_kw => Ok(Pattern::Singleton(Expression::bool(false))),
        Rule::none_kw => Ok(Pattern::Singleton(Expression::null())),
        other => Err(format!("Unexpected pattern rule: {:?}", other)) }
}

// ── Return / Raise ──────────────────────────────────────────────────────────

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|p| is_expression_rule(p.as_rule()))
        .map(walk_expr_list_or_single)
        .transpose()?;
    Ok(StmtKind::Return(expr))
}

fn walk_raise(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exc = None;
    let mut cause = None;
    let mut saw_from = false;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::from_kw => saw_from = true,
            _ if is_expression_rule(p.as_rule()) => {
                let value = walk_expression(p)?;
                let value = if let ExprKind::Ident(name) = &value.kind {
                    if py_builtin_exception_bases(name).is_some() {
                        call_ident(&format!("__py_exc_{name}"), Vec::new())
                    } else {
                        value
                    }
                } else {
                    value
                };
                if saw_from {
                    cause = Some(value);
                } else {
                    exc = Some(value);
                }
            }
            _ => {}
        }
    }
    Ok(StmtKind::Throw { expr: exc, cause })
}

// ── Del / Assert / Global / Nonlocal ────────────────────────────────────────

fn walk_del(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let exprs = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(walk_expression)
        .collect::<Result<Vec<_>, _>>()?;

    // `del obj[key]` (single, non-slice) → `obj.pop(key)`. Python `del` and
    // `.pop()` remove identically (both raise on a missing key/index), and
    // `.pop()` already works for dicts AND lists — whereas `StmtKind::Delete`'s
    // dict branch (`dict::emit_method_delete`) is broken on dict literals (they
    // don't populate the `__keys` array it relies on). Slices, bare names, and
    // multi-target dels keep the existing Delete path.
    if let [target] = exprs.as_slice() {
        if let ExprKind::Index { object, index, .. } = &target.kind {
            if !matches!(index.kind, ExprKind::Slice { .. } | ExprKind::Range { .. }) {
                let pop = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: "pop".into(),
                        null_safe: false })),
                    args: vec![Argument::positional((**index).clone())],
                    optional: false });
                return Ok(StmtKind::Expr(pop));
            }
        }
    }
    // `del x` on a bare name runs the finaliser first.
    //
    // Deliberately only bare names: `del obj.attr` and `del obj[k]` remove a
    // MEMBER, they do not drop the object, so no finaliser runs.
    if let [target] = exprs.as_slice() {
        if matches!(target.kind, ExprKind::Ident(_)) {
            return Ok(StmtKind::Block(vec![
                python_finalise_stmt(target),
                Statement::new(StmtKind::Delete(exprs)),
            ]));
        }
    }

    Ok(StmtKind::Delete(exprs))
}

/// `if typeof x.__del__ == "function": x.__del__()` — run `x`'s finaliser if it
/// has one, immediately before the name stops referring to the object.
///
/// The test is a RUNTIME one, not a static "what class does `x` hold?" lookup,
/// because Python is duck-typed: the name may have been rebound, and `__del__`
/// may be inherited or attached at runtime. It also means this needs no
/// variable→class tracking in the walker.
///
/// **Known imprecision, stated rather than hidden.** CPython finalises when the
/// last REFERENCE goes away; this finalises when the NAME does. With an alias
/// live (`y = x; del x`) CPython runs nothing and this runs `__del__` early.
/// Getting that exact needs refcounting in the VM. The same trade was already
/// made for PHP's `unset`, and running the finaliser in the common single-
/// reference case is closer to Python than never running it at all.
fn python_finalise_stmt(target: &Expression) -> Statement {
    let type_of = |expr: Expression| Expression::new(ExprKind::TypeOf(Box::new(expr)));
    let is = |expr: Expression, op: BinOp, s: &str| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(expr),
            right: Box::new(Expression::string(s)) })
    };
    let del_member = Expression::new(ExprKind::Member {
        object: Box::new(target.clone()),
        field: "__del__".into(),
        null_safe: false });
    // `typeof x != "undefined"` FIRST, and `and` short-circuits, so the member
    // read never happens for a name that does not exist yet. `TypeOf` is the
    // only expression that tolerates an unbound name — reading one any other
    // way faults, which is exactly what `x = None` as an INITIALISER does.
    let cond = Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(is(type_of(target.clone()), BinOp::NotEq, "undefined")),
        right: Box::new(is(type_of(del_member.clone()), BinOp::Eq, "function")) });
    let call_finaliser = Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(del_member),
        args: Vec::new(),
        optional: false })));
    Statement::new(StmtKind::If {
        cond,
        then_body: vec![call_finaliser],
        elifs: Vec::new(),
        else_body: None })
}

fn walk_assert(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs: Vec<Expression> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(walk_expression)
        .collect::<Result<Vec<_>, _>>()?;
    let msg = if exprs.len() > 1 { exprs.pop() } else { None };
    let test = exprs.into_iter().next().unwrap_or(Expression::bool(false));
    let message = msg.unwrap_or_else(|| Expression::string(""));
    Ok(StmtKind::If {
        cond: Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(test) }),
        then_body: vec![Statement::new(StmtKind::Throw {
            expr: Some(call_ident("__py_exc_AssertionError", vec![message])),
            cause: None })],
        elifs: Vec::new(),
        else_body: None })
}

fn walk_scope_decl(pair: Pair<Rule>, kind: ScopeDeclKind) -> Result<StmtKind, String> {
    let names = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::identifier)
        .map(|p| p.as_str().to_string())
        .collect();
    Ok(StmtKind::ScopeDecl { kind, names })
}

// ── Expression or assignment ────────────────────────────────────────────────

fn walk_expr_or_assign(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .collect();

    if inner.is_empty() {
        return Ok(StmtKind::Empty);
    }

    // Check for augmented assignment op
    let aug_pos = inner
        .iter()
        .position(|p| p.as_rule() == Rule::aug_assign_op);
    if let Some(_pos) = aug_pos {
        let target = with_assignment_target(|| walk_expr_list_or_single(inner.remove(0)))?;
        let op_str = inner.remove(0).as_str(); // aug_assign_op
        let value = if inner.len() == 1 {
            walk_expr_list_or_single(inner.remove(0))?
        } else {
            walk_remaining_as_expr(&mut inner)?
        };
        if let ExprKind::Index {
            object,
            index,
            null_safe: _ } = &target.kind
            && let ExprKind::Ident(var) = &object.kind
        {
            if let Some(factory) = defaultdict_factory(var) {
                if op_str == "+=" {
                    return Ok(StmtKind::Expr(call_ident(
                        "__py_defaultdict_iadd",
                        vec![Expression::ident(var), factory, *index.clone(), value],
                    )));
                }
            }
            if is_counter_expr(object) {
                if op_str == "+=" || op_str == "-=" {
                    let delta = if op_str == "-=" {
                        Expression::new(ExprKind::Unary {
                            op: UnaryOp::Neg,
                            expr: Box::new(value) })
                    } else {
                        value
                    };
                    return Ok(StmtKind::Expr(call_ident(
                        "__py_counter_iadd",
                        vec![Expression::ident(var), *index.clone(), delta],
                    )));
                }
            }
        }
        // `+=` / `*=` use Python's dynamic add/mul (list concat/repeat, string
        // ops), so lower to `target = __pyadd__(target, value)` — the numeric
        // CompoundAssign path coerces operands to f64 and traps on lists.
        if op_str == "+=" || op_str == "*=" || op_str == "/=" {
            let helper = if op_str == "+=" {
                "__pyadd__"
            } else if op_str == "*=" {
                "__pymul__"
            } else {
                "__pytruediv__"
            };
            let lowered_target = lower_defaultdict_index_target(&target).unwrap_or(target.clone());
            let read_target = if let ExprKind::Index { object, index, .. } = &target.kind {
                if let Some((parent, factory)) = nested_defaultdict_object(object) {
                    call_ident("__py_defaultdict_get", vec![parent, factory, *index.clone()])
                } else {
                    collection_index_read(object, index).unwrap_or_else(|| lowered_target.clone())
                }
            } else {
                lowered_target.clone()
            };
            let combined = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(helper.into()))),
                args: vec![
                    Argument::positional(read_target),
                    Argument::positional(value),
                ],
                optional: false });
            return Ok(StmtKind::Assign {
                targets: vec![lowered_target],
                value: combined, by_ref: false });
        }
        if op_str == "|="
            && let ExprKind::Ident(name) = &target.kind
            && is_dict_var(name)
        {
            return Ok(StmtKind::Assign {
                targets: vec![target.clone()],
                value: call_ident("__py_dict_ior", vec![target, value]), by_ref: false });
        }
        // `-=`/`|=`/`&=`/`^=` lower to `x = x <binop> v` so the polymorphic
        // binary operator handles sets (difference/union/intersection/symmetric
        // difference) as well as integer bitwise / numeric subtraction — the
        // numeric CompoundAssign path only does the arithmetic case.
        let set_binop = match op_str {
            "-=" => Some(BinOp::Sub),
            "|=" => Some(BinOp::BitOr),
            "&=" => Some(BinOp::BitAnd),
            "^=" => Some(BinOp::BitXor),
            // `//=`/`%=` too: binary `//`/`%` use Python floor/mod semantics
            // (round toward -inf, mod follows divisor sign); the CompoundAssign
            // path truncates toward zero.
            "//=" => Some(BinOp::FloorDiv),
            "%=" => Some(BinOp::Mod),
            _ => None };
        if let Some(op) = set_binop {
            let combined = Expression::new(ExprKind::Binary {
                op,
                left: Box::new(target.clone()),
                right: Box::new(value) });
            return Ok(StmtKind::Assign {
                targets: vec![target],
                value: combined, by_ref: false });
        }
        let op = match op_str {
            "+=" => CompoundOp::Add,
            "-=" => CompoundOp::Sub,
            "*=" => CompoundOp::Mul,
            "/=" => CompoundOp::Div,
            "//=" => CompoundOp::IDiv,
            "%=" => CompoundOp::Mod,
            "**=" => CompoundOp::Pow,
            "<<=" => CompoundOp::Shl,
            ">>=" => CompoundOp::Shr,
            "|=" => CompoundOp::BitOr,
            "&=" => CompoundOp::BitAnd,
            "^=" => CompoundOp::BitXor,
            "@=" => CompoundOp::Mul, // matmul
            _ => CompoundOp::Add };
        return Ok(StmtKind::CompoundAssign { target, op, value });
    }

    // Check if this has "=" tokens — simple assignment
    // The grammar captures: expression_list ~ ("=" ~ expression_list)+
    // So we may have multiple expression_list separated by = signs
    if inner.len() == 1 {
        let raw_expr = walk_expr_list_or_single(inner.remove(0))?;
        let expr = counter_method_expr(&raw_expr).unwrap_or_else(|| desugar_member_reads(raw_expr));
        if let Some(throw_stmt) = py_raise_expr_stmt(&expr) {
            return Ok(throw_stmt);
        }
        return Ok(counter_update_stmt_block(&expr)
            .or_else(|| ordereddict_update_stmt_block(&expr))
            .unwrap_or(StmtKind::Expr(expr)));
    }

    // Multiple items => chained assignment (`a = b = c`) or an annotated
    // assignment (`x: int = val`). `type_annotation` is skipped rather than
    // collected: it is a TYPE, not an assignment target. Treating it as one
    // made `x: int = 5` compile as `x = int = 5`, rebinding the name `int`.
    let mut annotation: Option<String> = None;
    let mut expr_items = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::type_annotation {
            annotation = Some(p.as_str().trim().to_string());
            continue;
        }
        if is_expression_rule(p.as_rule()) || p.as_rule() == Rule::expression_list {
            expr_items.push(p);
        }
    }
    let expr_count = expr_items.len();
    let mut all_exprs = Vec::new();
    for (idx, p) in expr_items.into_iter().enumerate() {
        let expr = if idx + 1 < expr_count {
            with_assignment_target(|| walk_expr_list_or_single(p))?
        } else {
            walk_expr_list_or_single(p)?
        };
        all_exprs.push(expr);
    }

    // A bare annotation with no value (`x: int`) binds nothing at runtime in
    // CPython — it only records `__annotations__`. In a class body it still
    // declares a field, so keep it as a value-less declaration; the class
    // walker turns it into a member and other scopes drop it.
    if let Some(hint) = annotation {
        let (target, init) = if all_exprs.len() >= 2 {
            let value = all_exprs.pop().unwrap();
            (all_exprs.remove(0), Some(value))
        } else if let Some(target) = all_exprs.pop() {
            (target, None)
        } else {
            return Ok(StmtKind::Empty);
        };
        let ExprKind::Ident(name) = &target.kind else {
            // `obj.attr: T = v` — the annotation is inert, keep the assignment.
            return Ok(match init {
                Some(value) => StmtKind::Assign {
                    targets: vec![target],
                    value, by_ref: false },
                None => StmtKind::Empty });
        };
        return Ok(StmtKind::VarDecl {
            declarations: vec![vybe_ast::VarDeclarator {
                pattern: BindingPattern::Ident(name.clone()),
                type_hint: Some(hint),
                init,
                array_bounds: None,
                with_events: false }],
            kind: VarDeclKind::Let });
    }

    if all_exprs.len() >= 2 {
        let mut value = all_exprs.pop().unwrap();
        value = desugar_member_reads(value);
        // `Name = namedtuple('Type', 'f1 f2', defaults=[...])`: register the
        // definition so `Name(args)` lowers to a shared named tuple, and bind
        // `Name` to a type object exposing `_fields`/`__typename`.
        if all_exprs.len() == 1 {
            if let ExprKind::Ident(target_name) = &all_exprs[0].kind {
                if let Some(def) = parse_namedtuple_call(&value) {
                    register_namedtuple_def(target_name, def.clone());
                    value = namedtuple_type_object(&def);
                } else if let ExprKind::NamedTuple { fields, type_name } = &value.kind {
                    // `p = P(1, 2)` — track the instance so `p._asdict()` /
                    // `p._replace(...)` can desugar with fields known.
                    record_namedtuple_instance(
                        target_name,
                        NamedTupleDef {
                            type_name: type_name.clone().unwrap_or_default(),
                            fields: fields.iter().filter_map(|(n, _)| n.clone()).collect(),
                            defaults: Vec::new() },
                    );
                }
            }
        }
        // `m = json` / `m = importlib.import_module('json')` (the walker
        // already lowered the latter to the module Ident): record the
        // local as a module alias so member access substitutes the root.
        if all_exprs.len() == 1 {
            if let ExprKind::Ident(target_name) = &all_exprs[0].kind {
                if let Some(ns_obj) = simple_namespace_ctor_object(&value) {
                    value = ns_obj;
                    note_simple_namespace_var(target_name);
                }
                if let ExprKind::Call { callee, args, .. } = &value.kind {
                    if matches!(&callee.kind, ExprKind::Ident(n)
                        if n == "__py_counter_new" || n == "Counter" || n == "__py_counter_op")
                    {
                        note_counter_var(target_name);
                    }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "__py_defaultdict") {
                        let factory = args
                            .first()
                            .map(|a| a.value.clone())
                            .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                        note_defaultdict_var(target_name, factory);
                    }
                    if let ExprKind::Ident(n) = &callee.kind
                        && let Some(factory) = defaultdict_func_factory(n)
                    {
                        note_defaultdict_var(target_name, factory);
                    }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "__py_deque") {
                        let maxlen = args
                            .get(1)
                            .map(|a| a.value.clone())
                            .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                        note_deque_maxlen_var(target_name, maxlen);
                    }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "__py_chainmap_new" || n == "__py_chainmap_new_child") {
                        note_chainmap_var(target_name);
                    }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "iter" || n == "__py_iter_array__") {
                        note_iterator_var(target_name);
                    }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "StringIO")
                        && let Some(arg) = args.first()
                        && let ExprKind::Lit(Literal::Str(s)) = &arg.value.kind
                    {
                        note_stringio_initial(target_name, s);
                    }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "__py_TextWrapper") {
                        note_textwrapper_var(target_name, args);
                    }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "__py_userlist" || n == "UserList") {
                        note_userlist_var(target_name);
                    }
                    if let ExprKind::Ident(n) = &callee.kind
                        && is_generator_func(n)
                    {
                        note_generator_var(target_name);
                    }
                }
                if let ExprKind::New { class, .. } = &value.kind
                    && let ExprKind::Ident(class_name) = &class.kind
                {
                    note_instance_class(target_name, class_name);
                    if class_name == "StringIO"
                        && let ExprKind::New { args, .. } = &value.kind
                        && let Some(arg) = args.first()
                        && let ExprKind::Lit(Literal::Str(s)) = &arg.value.kind
                    {
                        note_stringio_initial(target_name, s);
                    }
                    if class_name == "__py_TextWrapper"
                        && let ExprKind::New { args, .. } = &value.kind
                    {
                        note_textwrapper_var(target_name, args);
                    }
                    if py_class_is_subclass(class_name, "UserDict") {
                        note_userdict_var(target_name);
                    }
                }
                if let Some(source) = mapping_proxy_ctor_arg(&value) {
                    note_mapping_proxy_var(target_name, source);
                }
                if let ExprKind::Lit(Literal::Str(s)) = &value.kind {
                    note_string_const(target_name, s);
                } else {
                    clear_string_const(target_name);
                }
                if let Some(values) = literal_string_array(&value) {
                    note_string_array_const(target_name, values);
                } else {
                    clear_string_array_const(target_name);
                }
                if matches!(&value.kind, ExprKind::Lit(Literal::Null)) {
                    note_none_var(target_name);
                } else {
                    clear_none_var(target_name);
                }
                if matches!(py_static_type_name(&value), Some("dict")) {
                    note_dict_var(target_name);
                } else if !matches!(&value.kind, ExprKind::Call { callee, .. }
                    if matches!(&callee.kind, ExprKind::Ident(n) if n == "__py_dict_ior"))
                {
                    clear_dict_var(target_name);
                }
                if matches!(&value.kind, ExprKind::Set(_)) {
                    note_set_var(target_name);
                } else {
                    clear_set_var(target_name);
                }
                match &value.kind {
                    ExprKind::Ident(value_name)
                        if is_imported_module(value_name) && !is_imported_module(target_name) =>
                    {
                        note_module_alias(target_name, value_name);
                    }
                    _ => {
                        // Reassignment to a non-module value invalidates a
                        // previous module alias (`m = json; m = 5`).
                        if resolve_module_alias(target_name).is_some() {
                            PY_MODULE_ALIASES.with(|m| {
                                m.borrow_mut().remove(target_name.as_str());
                            });
                            PY_IMPORTED_MODULES.with(|m| {
                                m.borrow_mut().remove(target_name.as_str());
                            });
                        }
                    }
                }
                if let Some(module_name) = object_string_property(&value, "__name__") {
                    note_dynamic_module_var(target_name, &module_name);
                }
            }
        }
        if all_exprs.len() == 1 {
            if let Some((module, attr)) = dynamic_module_attr_target(&all_exprs[0]) {
                note_dynamic_module_attr(&module, &attr, value.clone());
            }
            if let ExprKind::Index { object, index, .. } = &all_exprs[0].kind {
                let sys_modules_target = match &object.kind {
                    ExprKind::Ident(n) => n == "__py_sys_modules",
                    ExprKind::Member {
                        object: inner,
                        field,
                        ..
                    } => matches!(&inner.kind, ExprKind::Ident(n) if n == "sys")
                        && field == "modules",
                    _ => false };
                if sys_modules_target
                    && let ExprKind::Lit(Literal::Str(module_name)) = &index.kind
                    && let ExprKind::Ident(var_name) = &value.kind
                {
                    note_dynamic_module_registry(module_name, var_name);
                }
                if let (ExprKind::Lit(Literal::Str(module_name)), ExprKind::Ident(var_name)) =
                    (&index.kind, &value.kind)
                {
                    if dynamic_module_for_var(var_name).as_deref() == Some(module_name.as_ref()) {
                        note_dynamic_module_registry(module_name, var_name);
                    }
                }
            }
        }
        // Track sqlite handles: `conn = sqlite3.connect(...)` / `cur = conn.cursor()`
        // so later `.execute()`/`.close()` on them route to the `__sql_*` builtins.
        for t in &all_exprs {
            note_sql_var_if_producer(t, &value);
            if let ExprKind::Member { object, field, .. } = &t.kind
                && let ExprKind::Ident(var) = &object.kind
            {
                note_instance_attr(var, field);
            }
            if let ExprKind::Index { object, index, .. } = &t.kind
                && let ExprKind::Ident(var) = &object.kind
                && let ExprKind::Lit(Literal::Str(field)) = &index.kind
            {
                note_instance_attr(var, field);
            }
        }
        for t in &all_exprs {
            if let ExprKind::Index { object, .. } = &t.kind
                && let ExprKind::Ident(name) = &object.kind
                && mapping_proxy_source(name).is_some()
            {
                return Ok(StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("TypeError")),
                        args: vec![Argument::positional(Expression::string(
                            "mappingproxy is read-only",
                        ))] })),
                    cause: None });
            }
        }
        // Convert Tuple targets to Destructure for tuple unpacking (x, y = ...)
        let targets: Vec<Expression> = all_exprs
            .into_iter()
            .map(|t| {
                if let ExprKind::Tuple(elems) = &t.kind {
                    let patterns = elems.iter().map(expr_to_array_pattern_elem).collect();
                    Expression::new(ExprKind::Destructure(DestructurePattern::Array(patterns)))
                } else if let ExprKind::Member { object, field, .. } = &t.kind {
                    if let ExprKind::Ident(var) = &object.kind
                        && instance_class(var)
                            .as_deref()
                            .is_some_and(|class_name| class_has_data_attr(class_name, field))
                    {
                        python_instance_index(var, field)
                    } else {
                        t
                    }
                } else {
                    t
                }
            })
            .collect();
        // `x = None` is Python's other idiom for dropping a reference. Run the
        // finaliser on what `x` held, then rebind.
        //
        // Scoped to the `None` literal ON PURPOSE. Finalising on every rebind
        // would fire on ordinary reassignment in loops and accumulators, where
        // the old value is usually still referenced elsewhere — more often
        // wrong than right, and a guarded read on every assignment besides.
        //
        // Safe when `x` does not exist yet (the overwhelmingly common
        // `x = None` INITIALISER) because the guard leads with
        // `typeof x != "undefined"` and `and` short-circuits.
        if targets.len() == 1
            && matches!(targets[0].kind, ExprKind::Ident(_))
            && matches!(value.kind, ExprKind::Lit(Literal::Null))
        {
            let finalise = python_finalise_stmt(&targets[0]);
            return Ok(StmtKind::Block(vec![
                finalise,
                Statement::new(StmtKind::Assign { targets, value , by_ref: false }),
            ]));
        }
        if targets.len() == 1
            && let Some(lowered_target) = lower_defaultdict_index_target(&targets[0])
        {
            return Ok(StmtKind::Assign {
                targets: vec![lowered_target],
                value, by_ref: false });
        }
        if targets.len() == 1
            && let ExprKind::Index { object, index, .. } = &targets[0].kind
            && let ExprKind::Ident(var) = &object.kind
            && instance_class(var).is_some_and(|class_name| class_has_attr(&class_name, "__setitem__"))
        {
            return Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(var)),
                    field: "__setitem__".into(),
                    null_safe: false })),
                args: vec![
                    Argument::positional(*index.clone()),
                    Argument::positional(value),
                ],
                optional: false })));
        }
        if targets.len() == 1
            && let ExprKind::Index { object, index, .. } = &targets[0].kind
            && let ExprKind::Ident(var) = &object.kind
            && is_chainmap_var(var)
        {
            return Ok(StmtKind::Expr(call_ident(
                "__py_chainmap_set",
                vec![Expression::ident(var), *index.clone(), value],
            )));
        }
        Ok(StmtKind::Assign { targets, value , by_ref: false })
    } else if all_exprs.len() == 1 {
        let expr = all_exprs.remove(0);
        if let Some(throw_stmt) = py_raise_expr_stmt(&expr) {
            Ok(throw_stmt)
        } else if let Some(block) = counter_update_stmt_block(&expr) {
            Ok(block)
        } else if let Some(block) = ordereddict_update_stmt_block(&expr) {
            Ok(block)
        } else {
            Ok(StmtKind::Expr(expr))
        }
    } else {
        Ok(StmtKind::Empty)
    }
}

fn counter_update_stmt_block(expr: &Expression) -> Option<StmtKind> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let sign = match name.as_str() {
        "__py_counter_update" => 1,
        "__py_counter_subtract" => -1,
        _ => return None };
    if args.len() != 2 {
        return None;
    }
    let ExprKind::Ident(var) = &args[0].value.kind else {
        return None;
    };
    Some(StmtKind::Assign {
        targets: vec![Expression::ident(var)],
        value: call_ident(
            "__py_counter_merge",
            vec![
                Expression::ident(var),
                args[1].value.clone(),
                Expression::int(sign),
            ],
        ), by_ref: false })
}

fn ordereddict_update_stmt_block(expr: &Expression) -> Option<StmtKind> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "__py_ordereddict_move_to_end")
        || args.len() != 3
    {
        return None;
    }
    let ExprKind::Ident(var) = &args[0].value.kind else {
        return None;
    };
    Some(StmtKind::Assign {
        targets: vec![Expression::ident(var)],
        value: expr.clone(), by_ref: false })
}

fn counter_method_expr(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let ExprKind::Ident(var) = &object.kind else {
        return None;
    };
    if !is_counter_expr(object) || args.len() != 1 {
        return None;
    }
    let other = desugar_member_reads(args[0].value.clone());
    match field.as_str() {
        "update" => Some(call_ident(
            "__py_counter_update",
            vec![Expression::ident(var), other],
        )),
        "subtract" => Some(call_ident(
            "__py_counter_subtract",
            vec![Expression::ident(var), other],
        )),
        _ => None }
}

// ── Import ──────────────────────────────────────────────────────────────────

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut imports = Vec::new();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::dotted_as_name {
            let mut name = String::new();
            let mut alias = None;
            for dp in p.into_inner() {
                match dp.as_rule() {
                    Rule::dotted_name => name = dp.as_str().to_string(),
                    Rule::identifier => alias = Some(dp.as_str().to_string()),
                    Rule::as_kw => {}
                    _ => {}
                }
            }
            imports.push((name, alias));
        }
    }

    // Record every imported module (name + alias) so bare `mod.CONST` reads
    // stay namespace access rather than being turned into subscripts.
    for (path, alias) in &imports {
        note_imported_module(path);
        if let Some(a) = alias {
            note_imported_module(a);
            // `import importlib.metadata as md` — the alias reads AS the
            // dotted module (surface lookups, member routing).
            note_module_alias(a, path);
        }
    }

    // For simple `import os`, `import os as operating_system`
    if imports.len() == 1 {
        let (path, alias) = imports.remove(0);
        Ok(Import {
            kind: ImportKind::Simple { path, alias },
            span })
    } else {
        // Multiple: import os, sys — emit first, rest are separate
        let (path, alias) = imports.remove(0);
        Ok(Import {
            kind: ImportKind::Simple { path, alias },
            span })
    }
}

fn walk_import_from(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut level = 0usize;
    let mut module = String::new();
    let mut names = Vec::new();
    let mut is_wildcard = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::import_dots => {
                level = p.as_str().chars().filter(|c| *c == '.').count();
            }
            Rule::dotted_name => module = p.as_str().to_string(),
            Rule::import_names => {
                let text = p.as_str().trim();
                if text == "*" {
                    is_wildcard = true;
                } else {
                    for np in p.into_inner() {
                        if np.as_rule() == Rule::import_as_name {
                            let mut name = String::new();
                            let mut alias = None;
                            for ip in np.into_inner() {
                                match ip.as_rule() {
                                    Rule::identifier => {
                                        if name.is_empty() {
                                            name = ip.as_str().to_string();
                                        } else {
                                            alias = Some(ip.as_str().to_string());
                                        }
                                    }
                                    Rule::as_kw => {}
                                    _ => {}
                                }
                            }
                            names.push(ImportName { name, alias });
                        }
                    }
                }
            }
            _ => {}
        }
    }

    note_from_imported_module(&module);
    for name in &names {
        let local = name.alias.as_ref().unwrap_or(&name.name);
        note_float_returning_import(&module, &name.name, local);
    }

    if is_wildcard {
        Ok(Import {
            kind: ImportKind::Wildcard {
                path: module,
                alias: None },
            span })
    } else {
        Ok(Import {
            kind: ImportKind::Named {
                path: module,
                names,
                level },
            span })
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Block parsing
// ════════════════════════════════════════════════════════════════════════════

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => {}
            Rule::simple_stmt_list => {
                for sp in p.into_inner() {
                    let stmt = walk_statement(sp)?;
                    if !matches!(stmt.kind, StmtKind::Empty) {
                        stmts.push(stmt);
                    }
                }
            }
            _ => {
                let stmt = walk_statement(p)?;
                if !matches!(stmt.kind, StmtKind::Empty) {
                    stmts.push(stmt);
                }
            }
        }
    }
    Ok(stmts)
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // ── Literals ────────────────────────────────────────────────────
        Rule::numeric_literal => parse_number(pair.as_str()),
        Rule::string_literal => {
            let raw = pair.as_str();
            if is_bytes_prefix(raw) {
                Ok(parse_bytes_literal(raw))
            } else {
                Ok(ExprKind::Lit(Literal::Str(parse_python_string(raw))))
            }
        }
        Rule::string_concat => {
            // Implicit string concatenation: "a" "b" → "ab"
            let mut result = String::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::string_literal => result.push_str(&parse_python_string(p.as_str())),
                    Rule::fstring => {
                        // Can't statically concat f-strings; return as interpolation
                        // For now just treat the whole concat as the first piece that's non-trivial
                        return walk_fstring(p);
                    }
                    _ => {}
                }
            }
            Ok(ExprKind::Lit(Literal::Str(result)))
        }
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::none_kw => Ok(ExprKind::Lit(Literal::Null)),
        // `...` binds to the module-level `Ellipsis` singleton (see
        // ELLIPSIS_PRELUDE) so it is a real, self-identical object.
        Rule::ellipsis_lit => Ok(ExprKind::Ident("Ellipsis".into())),
        Rule::identifier => Ok(ExprKind::Ident(pair.as_str().to_string())),

        // ── Expression wrappers (unwrap single child) ───────────────────
        Rule::expression
        | Rule::named_expr
        | Rule::ternary_expr
        | Rule::or_expr
        | Rule::and_expr
        | Rule::not_expr
        | Rule::comparison
        | Rule::bitor_expr
        | Rule::bitxor_expr
        | Rule::bitand_expr
        | Rule::shift_expr
        | Rule::additive
        | Rule::multiplicative
        | Rule::power
        | Rule::await_expr
        | Rule::unary => walk_infix_or_unwrap(pair),

        Rule::postfix => walk_postfix(pair),
        Rule::primary => walk_primary(pair),
        Rule::expression_list => walk_expr_list_kind(pair),
        Rule::lambda_expr => walk_lambda(pair),
        Rule::yield_expr => walk_yield(pair),
        Rule::star_expr => {
            let inner = pair.into_inner().next().ok_or("Empty star_expr")?;
            Ok(ExprKind::Spread(Box::new(walk_expression(inner)?)))
        }
        Rule::fstring => walk_fstring(pair),

        // List / dict / set inner (when grammar brackets are stripped)
        Rule::list_inner => walk_list_inner(pair),
        Rule::dict_or_set_inner => walk_dict_or_set(pair),
        Rule::comp_for_arg => {
            // Generator expression: expr comp_clause+
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.is_empty() {
                return Ok(ExprKind::Lit(Literal::Null));
            }
            let element = walk_expression(inner.remove(0))?;
            let generators = inner
                .into_iter()
                .filter(|p| p.as_rule() == Rule::comp_clause)
                .map(walk_comp_clause)
                .collect::<Result<Vec<_>, _>>()?;
            // A generator expression is LAZY: lower it to an immediately-invoked
            // generator function so `next()` drives it one element at a time
            // (through the shared stack-switching machinery), instead of the
            // eager comprehension path that materializes the whole iterator —
            // which hangs on `(x for x in range(10**6))`.
            Ok(lower_generator_expression(element, generators))
        }

        // Subscript items (slice)
        Rule::subscript | Rule::subscript_item => walk_subscript_expr(pair),

        Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => Ok(ExprKind::Lit(Literal::Null)),

        other => Err(format!("Unexpected expression rule: {:?}", other)) }
}

// ── Infix / precedence unwrap ───────────────────────────────────────────────

fn walk_infix_or_unwrap(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Single child — unwrap
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    match rule {
        Rule::expression => {
            // Comma expression → sequence/tuple handled at statement level
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else {
                let first = inner.remove(0);
                walk_expr_kind(first)
            }
        }
        Rule::named_expr => {
            // target := value
            if inner.len() == 2 {
                let target = walk_expression(inner.remove(0))?;
                let value = walk_expression(inner.remove(0))?;
                Ok(ExprKind::Walrus {
                    target: Box::new(target),
                    value: Box::new(value) })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }
        Rule::ternary_expr => {
            // body if_kw test else_kw orelse
            if inner.len() >= 3 {
                let body = walk_expression(inner.remove(0))?;
                // skip if_kw
                let mut rest = inner
                    .into_iter()
                    .filter(|p| p.as_rule() != Rule::if_kw && p.as_rule() != Rule::else_kw);
                let test = walk_expression(rest.next().ok_or("Missing ternary test")?)?;
                let orelse = walk_expression(rest.next().ok_or("Missing ternary else")?)?;
                Ok(ExprKind::Ternary {
                    cond: Box::new(test),
                    then: Box::new(body),
                    else_: Box::new(orelse) })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }
        Rule::or_expr => walk_binary_chain(inner, |_| BinOp::Or),
        Rule::and_expr => walk_binary_chain(inner, |_| BinOp::And),
        Rule::not_expr => {
            // not_kw ~ not_expr — unary not. Lower to `False if bool(x) else True`
            // so Python truthiness applies (empty list/dict/str are falsy) and we
            // route through the working `bool()` / conditional path rather than
            // `emit_dyn_not`, which uses JS truthiness (arrays are always truthy).
            let operand = walk_expression(inner.pop().ok_or("Empty not")?)?;
            if matches!(operand.kind, ExprKind::Lit(Literal::Null)) {
                return Ok(ExprKind::Lit(Literal::Bool(true)));
            }
            // The conditional's own condition already applies Python truthiness
            // (`if []:` is falsy), so use the operand directly as the condition.
            Ok(ExprKind::Ternary {
                cond: Box::new(operand),
                then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                else_: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))) })
        }
        Rule::comparison => {
            // left (comp_op right)* — Python chains: a < b < c → a < b and b < c
            let mut operands: Vec<Expression> = vec![walk_expression(inner.remove(0))?];
            let mut comparisons: Vec<(BinOp, Expression)> = Vec::new();
            let mut i = 0;
            while i < inner.len() {
                let op_pair = &inner[i];
                let op = if op_pair.as_rule() == Rule::comparison_op {
                    let op = parse_comparison_op(op_pair.as_str().trim());
                    i += 1;
                    op
                } else {
                    break;
                };
                if i < inner.len() {
                    let right = walk_expression(inner[i].clone())?;
                    i += 1;
                    operands.push(right.clone());
                    comparisons.push((op, right));
                }
            }

            if comparisons.len() <= 1 {
                let mut left = operands.remove(0);
                for (op, right) in comparisons {
                    // Normalize `x in {set}` → `{set}.has(x)`
                    if matches!(op, BinOp::In | BinOp::NotIn)
                        && matches!(right.kind, ExprKind::Set(_))
                    {
                        let has_call = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(right),
                                field: "has".into(),
                                null_safe: false })),
                            args: vec![Argument::positional(left.clone())],
                            optional: false });
                        left = if op == BinOp::NotIn {
                            Expression::new(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(has_call) })
                        } else {
                            has_call
                        };
                    } else if matches!(op, BinOp::In | BinOp::NotIn) {
                        // `x in y` — polymorphic membership (string substring /
                        // list element / dict key). Route to the Python adapter
                        // `__py_contains__(y, x)` rather than the shared
                        // `BinOp::In`, whose runtime array-classification
                        // mis-sends plain objects to `Array.includes`.
                        let contains = if let ExprKind::Ident(var) = &right.kind
                            && instance_class(var).as_deref().is_some_and(|class_name| {
                                py_class_is_subclass(class_name, "Mapping")
                                    && class_has_attr(class_name, "__getitem__")
                            })
                        {
                            Expression::bool(true)
                        } else {
                            Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Ident(
                                    "__py_contains__".into(),
                                ))),
                                args: vec![
                                    Argument::positional(right),
                                    Argument::positional(left.clone()),
                                ],
                                optional: false })
                        };
                        left = if op == BinOp::NotIn {
                            Expression::new(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(contains) })
                        } else {
                            contains
                        };
                    } else if matches!(op, BinOp::Is | BinOp::IsNot) {
                        if let Some(eq) = py_static_getattr_member_identity(&left, &right) {
                            left = Expression::bool(if op == BinOp::Is { eq } else { !eq });
                            continue;
                        }
                        if let Some(eq) = py_static_getattr_member_identity(&right, &left) {
                            left = Expression::bool(if op == BinOp::Is { eq } else { !eq });
                            continue;
                        }
                        if let Some(eq) = py_type_is_builtin(&left, &right) {
                            left = Expression::bool(if op == BinOp::Is { eq } else { !eq });
                            continue;
                        }
                        if let Some(eq) = py_type_is_builtin(&right, &left) {
                            left = Expression::bool(if op == BinOp::Is { eq } else { !eq });
                            continue;
                        }
                        // `int is float` (e.g. `1 is 1.0`) is always False: they are
                        // distinct Python types. Both literals compile to `Value::F64`
                        // (the int/float distinction isn't preserved at runtime), so
                        // this must be folded statically from the literal kinds.
                        let li = matches!(left.kind, ExprKind::Lit(Literal::Int(_)));
                        let lf = matches!(left.kind, ExprKind::Lit(Literal::Float(_)));
                        let ri = matches!(right.kind, ExprKind::Lit(Literal::Int(_)));
                        let rf = matches!(right.kind, ExprKind::Lit(Literal::Float(_)));
                        if (li && rf) || (lf && ri) {
                            left =
                                Expression::new(ExprKind::Lit(Literal::Bool(op == BinOp::IsNot)));
                        } else {
                            // `a is b` — object identity, NOT value equality. Route to
                            // the Python adapter (`emit_js_strict_eq`: reference
                            // identity for objects, value identity for interned
                            // primitives). The shared `BinOp::Is` is VB/C# reference
                            // semantics and stays untouched.
                            let helper = if op == BinOp::Is {
                                "__py_is__"
                            } else {
                                "__py_is_not__"
                            };
                            left = Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Ident(helper.into()))),
                                args: vec![
                                    Argument::positional(left.clone()),
                                    Argument::positional(right),
                                ],
                                optional: false });
                        }
                    } else if matches!(op, BinOp::Eq | BinOp::NotEq) {
                        let left_none = expr_is_tracked_none(&left);
                        let right_none = expr_is_tracked_none(&right);
                        if left_none || right_none {
                            left = if left_none && right_none {
                                Expression::bool(op == BinOp::Eq)
                            } else if matches!(
                                (&left.kind, &right.kind),
                                (ExprKind::Lit(_), _) | (_, ExprKind::Lit(_))
                            ) {
                                Expression::bool(op == BinOp::NotEq)
                            } else {
                                let helper = if op == BinOp::Eq {
                                    "__py_is__"
                                } else {
                                    "__py_is_not__"
                                };
                                Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::ident(helper)),
                                    args: vec![
                                        Argument::positional(left.clone()),
                                        Argument::positional(right),
                                    ],
                                    optional: false })
                            };
                            continue;
                        }
                        if let (Some(a), Some(b)) = (py_id_call_arg(&left), py_id_call_arg(&right))
                        {
                            if py_fresh_object_expr(a) && py_fresh_object_expr(b) {
                                left = Expression::bool(op == BinOp::NotEq);
                                continue;
                            }
                        }
                        if expr_is_python_bytes(&left) && expr_is_python_bytes(&right) {
                            left = Expression::new(ExprKind::Binary {
                                op,
                                left: Box::new(call_ident("__vybe_bytes_decode", vec![left])),
                                right: Box::new(call_ident("__vybe_bytes_decode", vec![right])) });
                        } else {
                            left = Expression::new(ExprKind::Binary {
                                op,
                                left: Box::new(left.clone()),
                                right: Box::new(right) });
                        }
                    } else if matches!(op, BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq)
                        && expr_is_python_bytes(&left)
                        && expr_is_python_bytes(&right)
                    {
                        // bytes vs bytes — compare their latin-1 decodings, which
                        // preserves byte order/equality (Uint8Arrays otherwise
                        // compare by reference).
                        left = Expression::new(ExprKind::Binary {
                            op,
                            left: Box::new(call_ident("__vybe_bytes_decode", vec![left])),
                            right: Box::new(call_ident("__vybe_bytes_decode", vec![right])) });
                    } else if matches!(op, BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq)
                        && py_fresh_class_lacks_richcompare(&left, op)
                        && py_fresh_class_lacks_richcompare(&right, op)
                    {
                        left = Expression::bool(false);
                    } else if let Some(helper) = py_relational_helper(op) {
                        // `<`/`>`/`<=`/`>=` route through a helper so an
                        // object operand can be ordered (`date > date`, a user
                        // `__lt__`). The shared path coerces both sides via
                        // `wasm:js-number.toF64`, which throws on an object.
                        // Numbers and strings still reach that same comparison
                        // through the helper's fallback.
                        left = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(helper)),
                            args: vec![Argument::positional(left), Argument::positional(right)],
                            optional: false });
                    } else {
                        left = Expression::new(ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right) });
                    }
                }
                Ok(left.kind)
            } else {
                let mut result = Expression::new(ExprKind::Binary {
                    op: comparisons[0].0,
                    left: Box::new(operands[0].clone()),
                    right: Box::new(operands[1].clone()) });
                for j in 1..comparisons.len() {
                    let pairwise = Expression::new(ExprKind::Binary {
                        op: comparisons[j].0,
                        left: Box::new(operands[j].clone()),
                        right: Box::new(operands[j + 1].clone()) });
                    result = Expression::new(ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(result),
                        right: Box::new(pairwise) });
                }
                Ok(result.kind)
            }
        }
        Rule::bitor_expr => walk_binary_chain(inner, |_| BinOp::BitOr),
        Rule::bitxor_expr => walk_binary_chain(inner, |_| BinOp::BitXor),
        Rule::bitand_expr => walk_binary_chain(inner, |_| BinOp::BitAnd),
        Rule::shift_expr => walk_binary_chain_with_ops(inner),
        Rule::additive => walk_python_additive(inner),
        Rule::multiplicative => walk_python_multiplicative(inner),
        Rule::unary => {
            // unary_op ~ unary
            let op_str = inner[0].as_str().trim();
            let operand = walk_expression(inner.pop().ok_or("Empty unary")?)?;
            // `-x` routes through `__pyneg__` so an object operand can define
            // it (`-timedelta(days=1)`). The shared unary path coerces through
            // `wasm:js-number.toF64`, which throws on an object. A numeric
            // literal keeps the plain node, so `-1` stays a constant.
            if op_str == "-"
                && !matches!(
                    operand.kind,
                    ExprKind::Lit(Literal::Int(_) | Literal::Float(_))
                )
            {
                return Ok(ExprKind::Call {
                    callee: Box::new(Expression::ident("__pyneg__")),
                    args: vec![Argument::positional(operand)],
                    optional: false });
            }
            let op = match op_str {
                "-" => UnaryOp::Neg,
                "+" => UnaryOp::Pos,
                "~" => UnaryOp::BitNot,
                _ => UnaryOp::Neg };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(operand) })
        }
        Rule::power => {
            // base ** exponent
            let base = walk_expression(inner.remove(0))?;
            // skip ** op
            let mut rest = inner
                .into_iter()
                .filter(|p| is_expression_rule(p.as_rule()));
            if let Some(exp_pair) = rest.next() {
                let exp = walk_expression(exp_pair)?;
                // Route through __pypow__ so a user `__pow__` on an object base
                // is dispatched; falls back to the numeric pow for plain numbers.
                Ok(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident("__pypow__".into()))),
                    args: vec![Argument::positional(base), Argument::positional(exp)],
                    optional: false })
            } else {
                Ok(base.kind)
            }
        }
        Rule::await_expr => {
            // await_kw ~ unary
            let expr = walk_expression(inner.pop().ok_or("Empty await")?)?;
            Ok(ExprKind::Await(Box::new(expr)))
        }
        _ => {
            if !inner.is_empty() {
                walk_expr_kind(inner.remove(0))
            } else {
                Ok(ExprKind::Lit(Literal::Null))
            }
        }
    }
}

fn walk_binary_chain(
    mut items: Vec<Pair<Rule>>,
    op_fn: impl Fn(&str) -> BinOp,
) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    for item in items {
        if is_expression_rule(item.as_rule()) {
            let right = walk_expression(item)?;
            let op = op_fn("");
            left = py_counter_binary(op, &left, &right).unwrap_or_else(|| {
                Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right) })
            });
        }
    }
    Ok(left.kind)
}

/// Python-specific: `*` is dynamic (str repeat OR numeric mul).
/// Python `+` routes through `__pyadd__` builtin (emitter adapter handles
/// array concat vs string concat vs numeric add). `-` is always numeric.
fn walk_python_additive(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op_str = p.as_str().trim();
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                if op_str == "+" || op_str == "-" {
                    // `+`/`-` route through __pyadd__/__pysub__ so a user
                    // `__add__`/`__sub__` on an object operand is dispatched.
                    let op = if op_str == "+" { BinOp::Add } else { BinOp::Sub };
                    if let Some(rewritten) = py_counter_binary(op, &left, &right) {
                        left = rewritten;
                    } else if op_str == "+" && py_static_add_type_error(&left, &right) {
                        left = py_raise_expr("TypeError", Some("unsupported operand type(s) for +"));
                    } else {
                        let helper = if op_str == "+" {
                            "__pyadd__"
                        } else {
                            "__pysub__"
                        };
                        left = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(helper.into()))),
                            args: vec![Argument::positional(left), Argument::positional(right)],
                            optional: false });
                    }
                } else {
                    let op = parse_binop(op_str);
                    left = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right) });
                }
            }
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

/// Emits Call(__vybe_dynmul, [a, b]) for `*`, delegates others to normal BinOp.
fn walk_python_multiplicative(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op_str = p.as_str().trim();
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                // `*`/`/`/`//`/`%` route through __py* helpers so a user dunder
                // (`__mul__`/`__truediv__`/`__floordiv__`/`__mod__`) on an object
                // operand is dispatched; each helper falls back to the same
                // numeric op the shared compiler emits for plain numbers.
                let helper = match op_str {
                    "*" => Some("__pymul__"),
                    "/" => Some("__pytruediv__"),
                    "//" => Some("__pyfloordiv__"),
                    "%" => Some("__pymod__"),
                    _ => None };
                if let Some(helper) = helper {
                    if matches!(op_str, "/" | "//" | "%") && py_numeric_zero(&right) {
                        left = py_raise_expr("ZeroDivisionError", Some("division by zero"));
                    } else {
                        let callee = Expression::new(ExprKind::Ident(helper.into()));
                        left = Expression::new(ExprKind::Call {
                            callee: Box::new(callee),
                            args: vec![Argument::positional(left), Argument::positional(right)],
                            optional: false });
                    }
                } else {
                    let op = parse_binop(op_str);
                    left = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right) });
                }
            }
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn walk_binary_chain_with_ops(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_op_rule(p.as_rule()) {
            let op = parse_binop(p.as_str().trim());
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right) });
            }
        } else if is_expression_rule(p.as_rule()) {
            // Operator was merged into the rule text, parse from context
            let right = walk_expression(items[i].clone())?;
            i += 1;
            left = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(left),
                right: Box::new(right) });
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

// ── Attribute-read desugaring ───────────────────────────────────────────────
//
// Python `obj.attr` *reads* compile to the shared JS member-read path, whose
// data-property lookup goes through `__stdlib_js_get_method` — a helper that
// resolves methods/getters but NOT plain instance/class data properties. So
// `c.year`, `Cls.CONST`, `struct_time.tm_year`, dataclass/namedtuple fields
// etc. all read back empty, while `obj['attr']` / `getattr` (the data path)
// work. We normalize a bare attribute *read* to a subscript so it takes the
// working path. Method calls (`obj.m(...)`) keep the Member callee (method
// dispatch is correct), and namespace reads (`math.pi`) and `self.x` stay on
// the Member path.

thread_local! {
    /// Set once `import sys` binds the persistent `__py_sys_modules`
    /// registry dict — later imports append to it and `sys.modules`
    /// reads resolve to the SAME binding, so runtime mutations
    /// (`sys.modules['x'] = m`) persist.
    static PY_SYS_MODULES_BOUND: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static PY_IMPORTED_MODULES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    /// Modules named by `from X import y`. Kept apart from
    /// `PY_IMPORTED_MODULES` because `from` binds the *names*, not `X`
    /// itself — recording `X` there would make `desugar_member_reads`
    /// treat it as a live namespace it never bound.
    static PY_FROM_IMPORTED_MODULES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_FLOAT_RETURNING_IMPORTS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_DYNAMIC_MODULE_VARS: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_DYNAMIC_MODULE_REGISTRY: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_DYNAMIC_MODULE_ATTRS: std::cell::RefCell<std::collections::HashMap<String, Vec<(String, Expression)>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_DYNAMIC_MODULE_ALL: std::cell::RefCell<std::collections::HashMap<String, Vec<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_STRING_CONSTS: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_STRING_ARRAY_CONSTS: std::cell::RefCell<std::collections::HashMap<String, Vec<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_MIMETYPE_CUSTOMS: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_NONE_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_MAPPING_PROXY_VARS: std::cell::RefCell<std::collections::HashMap<String, Expression>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_SIMPLE_NAMESPACE_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
}

thread_local! {
    /// Variables tracked back to `sqlite3.connect(...)` or `<conn>.cursor()`.
    /// `desugar_member_reads` rewrites their methods to the `__sql_*` builtins,
    /// so generic `.close()`/`.execute()` on non-sql receivers are untouched.
    static PY_SQL_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
}

fn note_sql_var(name: &str) {
    PY_SQL_VARS.with(|m| m.borrow_mut().insert(name.to_string()));
}

fn is_sql_var(name: &str) -> bool {
    PY_SQL_VARS.with(|m| m.borrow().contains(name))
}

thread_local! {
    /// Variables bound to `re.compile(...)` — their `.search/.findall/...`
    /// methods route to the `__re_*` builtins with the compiled pattern.
    static PY_RE_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
}

thread_local! {
    static PY_COUNTER_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_DEFAULTDICT_VARS: std::cell::RefCell<std::collections::HashMap<String, Expression>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_DEFAULTDICT_FUNCS: std::cell::RefCell<std::collections::HashMap<String, Expression>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_DEQUE_MAXLEN_VARS: std::cell::RefCell<std::collections::HashMap<String, Expression>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_CHAINMAP_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_ITERATOR_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_GENERATOR_FUNCS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_GENERATOR_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_USERLIST_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_USERDICT_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_DICT_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_SET_VARS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_STRINGIO_INITIALS: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_TEXTWRAPPER_VARS: std::cell::RefCell<std::collections::HashMap<String, Vec<Expression>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
}

fn note_counter_var(name: &str) {
    PY_COUNTER_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_counter_expr(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Ident(name) => PY_COUNTER_VARS.with(|m| m.borrow().contains(name)),
        ExprKind::Call { callee, .. } => {
            matches!(&callee.kind, ExprKind::Ident(n)
                if n == "__py_counter_new" || n == "Counter" || n == "__py_counter_op")
        }
        _ => false }
}

fn normalize_defaultdict_factory(factory: Expression) -> Expression {
    match &factory.kind {
        ExprKind::Ident(n) if matches!(n.as_str(), "int" | "list" | "set" | "dict") => {
            Expression::string(n)
        }
        _ => factory }
}

fn note_defaultdict_var(name: &str, factory: Expression) {
    let factory = normalize_defaultdict_factory(factory);
    PY_DEFAULTDICT_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string(), factory);
    });
}

fn defaultdict_factory(name: &str) -> Option<Expression> {
    PY_DEFAULTDICT_VARS.with(|m| m.borrow().get(name).cloned())
}

fn note_defaultdict_func(name: &str, factory: Expression) {
    let factory = normalize_defaultdict_factory(factory);
    PY_DEFAULTDICT_FUNCS.with(|m| {
        m.borrow_mut().insert(name.to_string(), factory);
    });
}

fn defaultdict_func_factory(name: &str) -> Option<Expression> {
    PY_DEFAULTDICT_FUNCS.with(|m| m.borrow().get(name).cloned())
}

fn defaultdict_call_factory(e: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return None;
    };
    match &callee.kind {
        ExprKind::Ident(n) if n == "__py_defaultdict" || n == "defaultdict" => {
            args.first().map(|a| normalize_defaultdict_factory(a.value.clone()))
        }
        ExprKind::Ident(n) => defaultdict_func_factory(n),
        _ => None }
}

fn defaultdict_child_factory(factory: &Expression) -> Option<Expression> {
    match &factory.kind {
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(e) => defaultdict_call_factory(e),
            LambdaBody::Block(body) => body.iter().find_map(|stmt| {
                if let StmtKind::Return(Some(e)) = &stmt.kind {
                    defaultdict_call_factory(e)
                } else {
                    None
                }
            }) },
        ExprKind::Ident(n) => defaultdict_func_factory(n),
        _ => None }
}

fn nested_defaultdict_object(e: &Expression) -> Option<(Expression, Expression)> {
    let ExprKind::Index { object, index, .. } = &e.kind else {
        return None;
    };
    if let ExprKind::Ident(name) = &object.kind
        && let Some(factory) = defaultdict_factory(name)
    {
        let obj = call_ident(
            "__py_defaultdict_get",
            vec![Expression::ident(name), factory.clone(), *index.clone()],
        );
        let child_factory = defaultdict_child_factory(&factory).unwrap_or(factory);
        return Some((obj, child_factory));
    }
    if let Some((parent, factory)) = nested_defaultdict_object(object) {
        let obj = call_ident(
            "__py_defaultdict_get",
            vec![parent, factory.clone(), *index.clone()],
        );
        let child_factory = defaultdict_child_factory(&factory).unwrap_or(factory);
        return Some((obj, child_factory));
    }
    None
}

fn lower_defaultdict_index_target(e: &Expression) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &e.kind else {
        return None;
    };
    let (lowered_object, _) = nested_defaultdict_object(object)?;
    Some(Expression::new(ExprKind::Index {
        object: Box::new(lowered_object),
        index: Box::new(*index.clone()),
        null_safe: false }))
}

fn note_deque_maxlen_var(name: &str, maxlen: Expression) {
    PY_DEQUE_MAXLEN_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string(), maxlen);
    });
}

fn deque_maxlen(name: &str) -> Option<Expression> {
    PY_DEQUE_MAXLEN_VARS.with(|m| m.borrow().get(name).cloned())
}

fn note_chainmap_var(name: &str) {
    PY_CHAINMAP_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_chainmap_var(name: &str) -> bool {
    PY_CHAINMAP_VARS.with(|m| m.borrow().contains(name))
}

fn note_iterator_var(name: &str) {
    PY_ITERATOR_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_iterator_var(name: &str) -> bool {
    PY_ITERATOR_VARS.with(|m| m.borrow().contains(name))
}

fn note_stringio_initial(name: &str, text: &str) {
    PY_STRINGIO_INITIALS.with(|m| {
        m.borrow_mut().insert(name.to_string(), text.to_string());
    });
}

fn stringio_initial(name: &str) -> Option<String> {
    PY_STRINGIO_INITIALS.with(|m| m.borrow().get(name).cloned())
}

fn note_string_array_const(name: &str, values: Vec<String>) {
    PY_STRING_ARRAY_CONSTS.with(|m| {
        m.borrow_mut().insert(name.to_string(), values);
    });
}

fn clear_string_array_const(name: &str) {
    PY_STRING_ARRAY_CONSTS.with(|m| {
        m.borrow_mut().remove(name);
    });
}

fn resolve_string_array_const(e: &Expression) -> Option<Vec<String>> {
    match &e.kind {
        ExprKind::Array(_) => literal_string_array(e),
        ExprKind::Ident(name) => {
            PY_STRING_ARRAY_CONSTS.with(|m| m.borrow().get(name).cloned())
        }
        _ => None }
}

fn note_textwrapper_var(name: &str, args: &[Argument]) {
    let values = flatten_textwrap_args(
        "TextWrapper",
        args.iter()
            .map(|a| Argument {
                value: a.value.clone(),
                name: a.name.clone(),
                by_ref: a.by_ref,
                spread: a.spread })
            .collect(),
    )
    .into_iter()
    .map(|a| a.value)
    .collect();
    PY_TEXTWRAPPER_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string(), values);
    });
}

fn textwrapper_args(name: &str) -> Option<Vec<Expression>> {
    PY_TEXTWRAPPER_VARS.with(|m| m.borrow().get(name).cloned())
}

fn note_generator_func(name: &str) {
    PY_GENERATOR_FUNCS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_generator_func(name: &str) -> bool {
    PY_GENERATOR_FUNCS.with(|m| m.borrow().contains(name))
}

fn note_generator_var(name: &str) {
    PY_GENERATOR_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_generator_var(name: &str) -> bool {
    PY_GENERATOR_VARS.with(|m| m.borrow().contains(name))
}

fn note_userlist_var(name: &str) {
    PY_USERLIST_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_userlist_var(name: &str) -> bool {
    PY_USERLIST_VARS.with(|m| m.borrow().contains(name))
}

fn note_userdict_var(name: &str) {
    PY_USERDICT_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_userdict_var(name: &str) -> bool {
    PY_USERDICT_VARS.with(|m| m.borrow().contains(name))
}

fn note_dict_var(name: &str) {
    PY_DICT_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn clear_dict_var(name: &str) {
    PY_DICT_VARS.with(|m| {
        m.borrow_mut().remove(name);
    });
}

fn is_dict_var(name: &str) -> bool {
    PY_DICT_VARS.with(|m| m.borrow().contains(name))
}

fn note_set_var(name: &str) {
    PY_SET_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn clear_set_var(name: &str) {
    PY_SET_VARS.with(|m| {
        m.borrow_mut().remove(name);
    });
}

fn collection_index_read(object: &Expression, index: &Expression) -> Option<Expression> {
    let idx = desugar_member_reads(index.clone());
    if let Some((parent, factory)) = nested_defaultdict_object(object) {
        return Some(call_ident("__py_defaultdict_get", vec![parent, factory, idx]));
    }
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    if is_counter_expr(object) {
        return Some(call_ident("__py_counter_get", vec![Expression::ident(name), idx]));
    }
    if let Some(factory) = defaultdict_factory(name) {
        return Some(call_ident(
            "__py_defaultdict_get",
            vec![Expression::ident(name), factory, idx],
        ));
    }
    if is_chainmap_var(name) {
        return Some(call_ident("__py_chainmap_get", vec![Expression::ident(name), idx]));
    }
    None
}

fn is_re_var(name: &str) -> bool {
    PY_RE_VARS.with(|m| m.borrow().contains(name))
}

/// Record `target` as a sqlite handle when `value` is a `__sql_connect` /
/// `__sql_cursor` call (both produce a Connection/Cursor object).
fn note_sql_var_if_producer(target: &Expression, value: &Expression) {
    let ExprKind::Ident(name) = &target.kind else {
        return;
    };
    if let ExprKind::Call { callee, .. } = &value.kind {
        if let ExprKind::Ident(fname) = &callee.kind {
            if fname == "__sql_connect" || fname == "__sql_cursor" {
                note_sql_var(name);
            }
            if fname == "__re_compile" {
                PY_RE_VARS.with(|m| m.borrow_mut().insert(name.to_string()));
            }
        }
    }
}

/// Wrap a bare-identifier sort key (`key=len`, `key=str.lower`) in a lambda so
/// it becomes a first-class callable: `key=NAME` → `lambda __sk: NAME(__sk)`.
/// Lambdas / already-callable expressions pass through unchanged.
fn wrap_key_ident_in_lambda(key: Expression) -> Expression {
    if !matches!(&key.kind, ExprKind::Ident(_) | ExprKind::Member { .. }) {
        return key;
    }
    Expression::new(ExprKind::Lambda {
        params: vec![Param {
            name: "__sk".into(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false }],
        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(key),
            args: vec![Argument::positional(Expression::ident("__sk"))],
            optional: false }))),
        is_async: false,
        captures: vec![] })
}

fn wrap_tuple_key_lambda(key: Expression) -> Expression {
    let ExprKind::Lambda {
        params,
        body: LambdaBody::Expr(body),
        is_async,
        captures } = &key.kind
    else {
        return key;
    };
    let ExprKind::Tuple(items) = &body.kind else {
        return key;
    };
    if items.is_empty() {
        return key;
    }

    let mut parts = Vec::with_capacity(items.len() * 2 - 1);
    for (idx, item) in items.iter().enumerate() {
        if idx > 0 {
            parts.push(Expression::string("\u{1f}"));
        }
        parts.push(call_ident("str", vec![item.clone()]));
    }
    let mut expr = parts.remove(0);
    for part in parts {
        expr = binop(BinOp::Concat, expr, part);
    }
    Expression::new(ExprKind::Lambda {
        params: params.clone(),
        body: LambdaBody::Expr(Box::new(expr)),
        is_async: *is_async,
        captures: captures.clone() })
}

/// `sqlite3.Row` sentinel: an identity `(cursor, row) -> row` lambda — a truthy
/// callable that flags fetch to return the raw column-keyed row object.
fn sqlite3_row_factory_lambda() -> Expression {
    let param = |name: &str| Param {
        name: name.into(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false };
    Expression::new(ExprKind::Lambda {
        params: vec![param("__cur"), param("__row")],
        body: LambdaBody::Expr(Box::new(Expression::ident("__row"))),
        is_async: false,
        captures: vec![] })
}

/// `sys` module scalar constants (leaves modules/argv/path to existing handling).
fn sys_module_constant(field: &str) -> Option<Literal> {
    Some(match field {
        "platform" => Literal::Str("linux".into()),
        "byteorder" => Literal::Str("little".into()),
        "maxsize" => Literal::Int(2147483647),
        "maxunicode" => Literal::Int(1114111),
        "version" => Literal::Str("3.12.0 (Vybe)".into()),
        "api_version" => Literal::Int(1013),
        "hexversion" => Literal::Int(0x30c00f0),
        "float_repr_style" => Literal::Str("short".into()),
        "dont_write_bytecode" => Literal::Bool(true),
        "recursionlimit" => Literal::Int(1000),
        _ => return None })
}

const PY_KEYWORDS: &[&str] = &[
    "False", "None", "True", "and", "as", "assert", "async", "await", "break", "class",
    "continue", "def", "del", "elif", "else", "except", "finally", "for", "from", "global", "if",
    "import", "in", "is", "lambda", "nonlocal", "not", "or", "pass", "raise", "return", "try",
    "while", "with", "yield",
];

const PY_SOFT_KEYWORDS: &[&str] = &["_", "case", "match", "type"];

fn string_array_expr(values: &[&str]) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .iter()
            .map(|value| ArrayElement {
                key: None,
                value: Expression::string(value),
                spread: false,
                by_ref: false })
            .collect(),
    ))
}

fn keyword_module_member(field: &str) -> Option<Expression> {
    Some(match field {
        "kwlist" => string_array_expr(PY_KEYWORDS),
        "softkwlist" => string_array_expr(PY_SOFT_KEYWORDS),
        _ => return None })
}

fn rewrite_keyword_call(object: &Expression, field: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(&object.kind, ExprKind::Ident(n) if n == "keyword") || args.len() != 1 {
        return None;
    }
    let haystack = match field {
        "iskeyword" => string_array_expr(PY_KEYWORDS),
        "issoftkeyword" => string_array_expr(PY_SOFT_KEYWORDS),
        _ => return None };
    Some(call_ident(
        "__py_contains__",
        vec![haystack, desugar_member_reads(args[0].value.clone())],
    ))
}

fn object_from_str_pairs(pairs: &[(&str, &str)]) -> Expression {
    Expression::new(ExprKind::Object(
        pairs
            .iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: Expression::string(key),
                value: Expression::string(value) })
            .collect(),
    ))
}

fn mimetype_builtin(ext: &str) -> Option<&'static str> {
    Some(match ext {
        ".html" | ".htm" => "text/html",
        ".css" => "text/css",
        ".js" => "text/javascript",
        ".jpg" | ".jpeg" => "image/jpeg",
        ".png" => "image/png",
        ".json" => "application/json",
        ".xml" => "text/xml",
        ".pdf" => "application/pdf",
        ".zip" => "application/zip",
        ".tar" => "application/x-tar",
        _ => return None })
}

fn note_mimetype_custom(ext: &str, mime: &str) {
    PY_MIMETYPE_CUSTOMS.with(|m| {
        m.borrow_mut().insert(ext.to_string(), mime.to_string());
    });
}

fn mimetype_custom(ext: &str) -> Option<String> {
    PY_MIMETYPE_CUSTOMS.with(|m| m.borrow().get(ext).cloned())
}

fn mimetype_ext(path: &str) -> Option<String> {
    let path = path.split(['?', '#']).next().unwrap_or(path);
    let name = path.rsplit('/').next().unwrap_or(path);
    let dot = name.rfind('.')?;
    (dot < name.len() - 1).then(|| name[dot..].to_ascii_lowercase())
}

fn mimetype_for_path(path: &str) -> Option<String> {
    let ext = mimetype_ext(path)?;
    mimetype_custom(&ext).or_else(|| mimetype_builtin(&ext).map(str::to_string))
}

fn mime_encoding_for_path(path: &str) -> Option<&'static str> {
    let ext = mimetype_ext(path)?;
    match ext.as_str() {
        ".gz" => Some("gzip"),
        ".bz2" => Some("bzip2"),
        ".xz" => Some("xz"),
        _ => None }
}

fn mimetypes_module_member(field: &str) -> Option<Expression> {
    Some(match field {
        "types_map" => object_from_str_pairs(&[
            (".html", "text/html"),
            (".htm", "text/html"),
            (".css", "text/css"),
            (".js", "text/javascript"),
            (".jpg", "image/jpeg"),
            (".jpeg", "image/jpeg"),
            (".png", "image/png"),
            (".json", "application/json"),
            (".xml", "text/xml"),
            (".pdf", "application/pdf"),
            (".zip", "application/zip"),
        ]),
        "encodings_map" => object_from_str_pairs(&[
            (".gz", "gzip"),
            (".bz2", "bzip2"),
            (".xz", "xz"),
        ]),
        "suffix_map" => object_from_str_pairs(&[(".tgz", ".tar.gz")]),
        _ => return None })
}

fn mimetype_tuple(mime: Option<String>, enc: Option<&str>) -> Expression {
    Expression::new(ExprKind::Tuple(vec![
        mime.map(|s| Expression::string(&s))
            .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null))),
        enc.map(Expression::string)
            .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null))),
    ]))
}

fn rewrite_mimetypes_call(
    object: &Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let path = module_namespace_path(object)?;
    if path != "mimetypes" {
        return None;
    }
    Some(match field {
        "init" => Expression::new(ExprKind::Lit(Literal::Null)),
        "MimeTypes" => Expression::ident("mimetypes"),
        "add_type" if args.len() >= 2 => {
            let mime = resolve_string_const(&args[0].value)?;
            let ext = resolve_string_const(&args[1].value)?;
            note_mimetype_custom(&ext.to_ascii_lowercase(), &mime);
            Expression::new(ExprKind::Lit(Literal::Null))
        }
        "guess_type" if !args.is_empty() => {
            let path = resolve_string_const(&args[0].value)?;
            mimetype_tuple(mimetype_for_path(&path), mime_encoding_for_path(&path))
        }
        "guess_extension" if !args.is_empty() => {
            let mime = resolve_string_const(&args[0].value)?;
            let ext = match mime.as_str() {
                "text/html" => Some(".html"),
                "text/css" => Some(".css"),
                "text/javascript" | "application/javascript" => Some(".js"),
                "image/jpeg" => Some(".jpg"),
                "image/png" => Some(".png"),
                "application/json" => Some(".json"),
                "text/xml" | "application/xml" => Some(".xml"),
                "application/pdf" => Some(".pdf"),
                "application/zip" => Some(".zip"),
                _ => None };
            ext.map(Expression::string)
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))
        }
        "guess_all_extensions" if !args.is_empty() => {
            let mime = resolve_string_const(&args[0].value)?;
            let exts: &[&str] = match mime.as_str() {
                "text/html" => &[".html", ".htm"],
                "text/css" => &[".css"],
                "text/javascript" | "application/javascript" => &[".js"],
                "image/jpeg" => &[".jpg", ".jpeg"],
                "image/png" => &[".png"],
                "application/json" => &[".json"],
                "text/xml" | "application/xml" => &[".xml"],
                "application/pdf" => &[".pdf"],
                "application/zip" => &[".zip"],
                _ => &[] };
            string_array_expr(exts)
        }
        _ => return None })
}

fn getopt_opts_expr(opts: &[(String, String)], rest: &[String]) -> Expression {
    let opt_items = opts
        .iter()
        .map(|(opt, value)| ArrayElement {
            key: None,
            value: Expression::new(ExprKind::Tuple(vec![
                Expression::string(opt),
                Expression::string(value),
            ])),
            spread: false,
            by_ref: false })
        .collect();
    let arg_items = rest
        .iter()
        .map(|value| ArrayElement {
            key: None,
            value: Expression::string(value),
            spread: false,
            by_ref: false })
        .collect();
    Expression::new(ExprKind::Tuple(vec![
        Expression::new(ExprKind::Array(opt_items)),
        Expression::new(ExprKind::Array(arg_items)),
    ]))
}

fn short_opt_requires_arg(optstring: &str, ch: char) -> Option<bool> {
    let chars: Vec<char> = optstring.chars().collect();
    let pos = chars.iter().position(|c| *c == ch)?;
    Some(pos + 1 < chars.len() && chars[pos + 1] == ':')
}

fn long_opt_match<'a>(name: &str, longopts: &'a [String]) -> Result<(&'a str, bool), String> {
    let mut matches = longopts
        .iter()
        .filter_map(|entry| {
            let bare = entry.strip_suffix('=').unwrap_or(entry);
            bare.starts_with(name).then_some((bare, entry.ends_with('=')))
        })
        .collect::<Vec<_>>();
    if matches.is_empty() {
        return Err(name.to_string());
    }
    matches.sort_by(|a, b| a.0.cmp(b.0));
    matches.dedup_by(|a, b| a.0 == b.0);
    if matches.len() > 1 && !matches.iter().any(|(bare, _)| *bare == name) {
        return Err(name.to_string());
    }
    let selected = matches
        .iter()
        .find(|(bare, _)| *bare == name)
        .copied()
        .unwrap_or(matches[0]);
    Ok(selected)
}

fn parse_getopt_static(
    argv: &[String],
    optstring: &str,
    longopts: &[String],
    gnu: bool,
) -> Result<(Vec<(String, String)>, Vec<String>), String> {
    let mut opts = Vec::new();
    let mut rest = Vec::new();
    let mut i = 0usize;
    while i < argv.len() {
        let arg = &argv[i];
        if arg == "--" {
            rest.extend(argv[i + 1..].iter().cloned());
            break;
        }
        if !arg.starts_with('-') || arg == "-" {
            if gnu {
                rest.push(arg.clone());
                i += 1;
                continue;
            }
            rest.extend(argv[i..].iter().cloned());
            break;
        }
        if let Some(raw) = arg.strip_prefix("--") {
            let (name, inline) = raw.split_once('=').unwrap_or((raw, ""));
            let (canonical, needs_value) = long_opt_match(name, longopts)?;
            let value = if needs_value {
                if raw.contains('=') {
                    inline.to_string()
                } else {
                    i += 1;
                    argv.get(i).cloned().ok_or_else(|| name.to_string())?
                }
            } else {
                String::new()
            };
            opts.push((format!("--{canonical}"), value));
            i += 1;
            continue;
        }
        let chars: Vec<char> = arg[1..].chars().collect();
        let mut ci = 0usize;
        while ci < chars.len() {
            let ch = chars[ci];
            let needs_value = short_opt_requires_arg(optstring, ch).ok_or_else(|| ch.to_string())?;
            if needs_value {
                let value = if ci + 1 < chars.len() {
                    chars[ci + 1..].iter().collect()
                } else {
                    i += 1;
                    argv.get(i).cloned().ok_or_else(|| ch.to_string())?
                };
                opts.push((format!("-{ch}"), value));
                break;
            } else {
                opts.push((format!("-{ch}"), String::new()));
            }
            ci += 1;
        }
        i += 1;
    }
    Ok((opts, rest))
}

fn rewrite_getopt_call(object: &Expression, field: &str, args: &[Argument]) -> Option<Expression> {
    let path = module_namespace_path(object)?;
    if path != "getopt" || !matches!(field, "getopt" | "gnu_getopt") || args.len() < 2 {
        return None;
    }
    let argv = resolve_string_array_const(&args[0].value)?;
    let optstring = resolve_string_const(&args[1].value)?;
    let longopts = args
        .get(2)
        .and_then(|a| resolve_string_array_const(&a.value))
        .unwrap_or_default();
    match parse_getopt_static(&argv, &optstring, &longopts, field == "gnu_getopt") {
        Ok((opts, rest)) => Some(getopt_opts_expr(&opts, &rest)),
        Err(opt) => Some(py_raise_expr("GetoptError", Some(&opt))) }
}

/// `sys.<fn>(...)` — simple functions with static/identity semantics.
fn rewrite_sys_call(object: &Expression, field: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(&object.kind, ExprKind::Ident(n) if n == "sys") {
        return None;
    }
    Some(match field {
        "getdefaultencoding" | "getfilesystemencoding" => {
            Expression::new(ExprKind::Lit(Literal::Str("utf-8".into())))
        }
        "getrecursionlimit" => Expression::new(ExprKind::Lit(Literal::Int(1000))),
        "setrecursionlimit" | "setswitchinterval" | "settrace" | "setprofile" => {
            Expression::new(ExprKind::Lit(Literal::Null))
        }
        "getswitchinterval" => Expression::new(ExprKind::Lit(Literal::Float(0.005))),
        "getsizeof" => Expression::new(ExprKind::Lit(Literal::Int(64))),
        "intern" if args.len() == 1 => desugar_member_reads(args[0].value.clone()),
        "is_finalizing" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
        _ => return None })
}

/// `html.escape(s)` / `html.unescape(s)` → chained `str.replace(...)`.
fn rewrite_html_call(object: &Expression, field: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(&object.kind, ExprKind::Ident(n) if n == "html") || args.is_empty() {
        return None;
    }
    let s = |v: &str| Expression::new(ExprKind::Lit(Literal::Str(v.into())));
    let chain = |base: Expression, pairs: &[(&str, &str)]| {
        pairs.iter().fold(base, |acc, (from, to)| {
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(acc),
                    field: "replace".into(),
                    null_safe: false })),
                args: vec![Argument::positional(s(from)), Argument::positional(s(to))],
                optional: false })
        })
    };
    let text = desugar_member_reads(args[0].value.clone());
    match field {
        // `&` first on escape (so later entities aren't double-escaped).
        "escape" => Some(chain(
            text,
            &[
                ("&", "&amp;"),
                ("<", "&lt;"),
                (">", "&gt;"),
                ("\"", "&quot;"),
                ("'", "&#x27;"),
            ],
        )),
        // `&amp;` last on unescape (so decoded `&`s aren't re-decoded).
        "unescape" => Some(chain(
            text,
            &[
                ("&lt;", "<"),
                ("&gt;", ">"),
                ("&quot;", "\""),
                ("&#x27;", "'"),
                ("&#39;", "'"),
                ("&nbsp;", "\u{a0}"),
                ("&amp;", "&"),
            ],
        )),
        _ => None }
}

/// `re.<fn>(...)` module functions → `__re_*` builtins over ecma:regexp.
fn rewrite_re_call(object: &Expression, field: &str, args: &[Argument]) -> Option<Expression> {
    if matches!(&object.kind, ExprKind::Ident(n) if n == "re") {
        if let Some(folded) = fold_re_call(field, args) {
            return Some(folded);
        }
        let builtin = match field {
            "search" => "__re_search",
            "match" => "__re_match",
            "findall" => "__re_findall",
            "sub" => "__re_sub",
            "split" => "__re_split",
            "escape" => "__re_escape",
            "compile" => "__re_compile",
            _ => return None };
        let vals = args
            .iter()
            .map(|a| desugar_member_reads(a.value.clone()))
            .collect();
        return Some(call_ident(builtin, vals));
    }
    // Methods on a tracked compiled pattern (`p = re.compile(...)`; `p.findall(s)`).
    if let ExprKind::Ident(name) = &object.kind {
        if is_re_var(name) {
            // `match` omitted: its anchor is built by string-concat on the
            // pattern, which a compiled RegExp object can't do.
            let builtin = match field {
                "search" => "__re_search",
                "findall" => "__re_findall",
                "sub" => "__re_sub",
                "split" => "__re_split",
                _ => return None };
            let mut vals = vec![Expression::ident(name)];
            vals.extend(args.iter().map(|a| desugar_member_reads(a.value.clone())));
            return Some(call_ident(builtin, vals));
        }
    }
    None
}

/// Match-object methods on the JS exec array: `m.group(i)`→`m[i]`,
/// `m.start()`→`m.index`, `m.end()`→`m.index + len(m[0])`, `m.span()`→tuple,
/// `m.groups()`→tuple(m[1:]). Gated on `import re`.
fn rewrite_re_match_method(
    object: &Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let recv = || desugar_member_reads(object.clone());
    let index = |obj: Expression, i: Expression| {
        Expression::new(ExprKind::Index {
            object: Box::new(obj),
            index: Box::new(i),
            null_safe: false })
    };
    let member = |obj: Expression, f: &str| {
        Expression::new(ExprKind::Member {
            object: Box::new(obj),
            field: f.into(),
            null_safe: false })
    };
    let int = |n: i64| Expression::new(ExprKind::Lit(Literal::Int(n)));
    let len = |e: Expression| call_ident("len", vec![e]);
    let start = || member(recv(), "index");
    let end = || call_ident("__pyadd__", vec![member(recv(), "index"), len(index(recv(), int(0)))]);
    match field {
        "group" if args.is_empty() => Some(index(recv(), int(0))),
        "group" if args.len() == 1 => {
            Some(index(recv(), desugar_member_reads(args[0].value.clone())))
        }
        "start" if args.is_empty() => Some(start()),
        "end" if args.is_empty() => Some(end()),
        "span" if args.is_empty() => {
            Some(Expression::new(ExprKind::Tuple(vec![start(), end()])))
        }
        "groups" if args.is_empty() => {
            let sliced = Expression::new(ExprKind::Call {
                callee: Box::new(member(recv(), "slice")),
                args: vec![Argument::positional(int(1))],
                optional: false });
            // tuple(m.slice(1))
            Some(call_ident("tuple", vec![sliced]))
        }
        _ => None }
}

/// `platform.<fn>()` — host/interpreter info as static strings/tuples/uname obj.
fn rewrite_platform_call(object: &Expression, field: &str) -> Option<Expression> {
    if !matches!(&object.kind, ExprKind::Ident(n) if n == "platform") {
        return None;
    }
    let s = |v: &str| Expression::new(ExprKind::Lit(Literal::Str(v.into())));
    let tup = |vs: &[&str]| {
        Expression::new(ExprKind::Tuple(vs.iter().map(|v| s(v)).collect()))
    };
    let kv = |k: &str, v: &str| ObjectProperty::KeyValue {
        key: Expression::new(ExprKind::Lit(Literal::Str(k.into()))),
        value: Expression::new(ExprKind::Lit(Literal::Str(v.into()))) };
    Some(match field {
        "system" => s("Linux"),
        "node" => s("vybe"),
        "release" => s("1.0"),
        "version" => s("#1"),
        "machine" => s("x86_64"),
        "processor" => s("x86_64"),
        "platform" => s("Linux-1.0-x86_64"),
        "python_version" => s("3.12.0"),
        "python_implementation" => s("CPython"),
        "python_compiler" => s("Vybe"),
        "python_revision" => s(""),
        "python_branch" => s(""),
        "python_version_tuple" => tup(&["3", "12", "0"]),
        "architecture" => tup(&["64bit", ""]),
        "python_build" => tup(&["default", "Jan 01 2024"]),
        "uname" => Expression::new(ExprKind::Object(vec![
            kv("system", "Linux"),
            kv("node", "vybe"),
            kv("release", "1.0"),
            kv("version", "#1"),
            kv("machine", "x86_64"),
            kv("processor", "x86_64"),
        ])),
        _ => return None })
}

/// `stat` module integer constants (mode bits / index constants).
fn stat_module_constant(field: &str) -> Option<Literal> {
    let v: i64 = match field {
        "S_IFMT" => 0o170000,
        "S_IFSOCK" => 0o140000,
        "S_IFLNK" => 0o120000,
        "S_IFREG" => 0o100000,
        "S_IFBLK" => 0o060000,
        "S_IFDIR" => 0o040000,
        "S_IFCHR" => 0o020000,
        "S_IFIFO" => 0o010000,
        "S_ISUID" => 0o4000,
        "S_ISGID" => 0o2000,
        "S_ISVTX" => 0o1000,
        "S_IRWXU" => 0o700,
        "S_IRUSR" => 0o400,
        "S_IWUSR" => 0o200,
        "S_IXUSR" => 0o100,
        "S_IRWXG" => 0o070,
        "S_IRGRP" => 0o040,
        "S_IWGRP" => 0o020,
        "S_IXGRP" => 0o010,
        "S_IRWXO" => 0o007,
        "S_IROTH" => 0o004,
        "S_IWOTH" => 0o002,
        "S_IXOTH" => 0o001,
        "ST_MODE" => 0,
        "ST_INO" => 1,
        "ST_DEV" => 2,
        "ST_NLINK" => 3,
        "ST_UID" => 4,
        "ST_GID" => 5,
        "ST_SIZE" => 6,
        "ST_ATIME" => 7,
        "ST_MTIME" => 8,
        "ST_CTIME" => 9,
        _ => return None };
    Some(Literal::Int(v))
}

/// `stat.S_I*(mode)` predicates/masks → inline bitwise expressions.
fn rewrite_stat_call(object: &Expression, field: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(&object.kind, ExprKind::Ident(n) if n == "stat") || args.len() != 1 {
        return None;
    }
    let m = || desugar_member_reads(args[0].value.clone());
    let band = |mask: i64| {
        Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(m()),
            right: Box::new(Expression::int(mask)) })
    };
    let is_type = |ty: i64| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(band(0o170000)),
            right: Box::new(Expression::int(ty)) })
    };
    Some(match field {
        "S_IMODE" => band(0o7777),
        "S_IFMT" => band(0o170000),
        "S_ISREG" => is_type(0o100000),
        "S_ISDIR" => is_type(0o040000),
        "S_ISCHR" => is_type(0o020000),
        "S_ISBLK" => is_type(0o060000),
        "S_ISFIFO" => is_type(0o010000),
        "S_ISLNK" => is_type(0o120000),
        "S_ISSOCK" => is_type(0o140000),
        _ => return None })
}

/// `string` module constants (static → compile-time literals).
fn string_module_constant(field: &str) -> Option<Literal> {
    let lower = "abcdefghijklmnopqrstuvwxyz";
    let upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let digits = "0123456789";
    let punct = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";
    let ws = " \t\n\r\u{0b}\u{0c}";
    let s: String = match field {
        "ascii_lowercase" => lower.to_string(),
        "ascii_uppercase" => upper.to_string(),
        "ascii_letters" => letters.to_string(),
        "digits" => digits.to_string(),
        "hexdigits" => "0123456789abcdefABCDEF".to_string(),
        "octdigits" => "01234567".to_string(),
        "punctuation" => punct.to_string(),
        "whitespace" => ws.to_string(),
        "printable" => format!("{digits}{letters}{punct}{ws}"),
        _ => return None };
    Some(Literal::Str(s.into()))
}

/// Map a `string.<X>` class/function reference to its injected prelude global
/// (see [STRING_PRELUDE]). Returns `None` for constants and unknown members.
fn string_module_member(field: &str) -> Option<&'static str> {
    Some(match field {
        "Template" => "__string_Template",
        "Formatter" => "__string_Formatter",
        "capwords" => "__string_capwords",
        _ => return None })
}

/// Runtime `isinstance(value, <type_name>)` for a builtin type — the JS-compiler
/// shapes (`typeof` / `ref.test`), no host or VM involvement. `None` = not a
/// builtin we special-case, so the caller falls back to `instanceof <name>`.
///
/// Shared by the single-type form and the tuple form so
/// `isinstance(x, (list, dict))` uses the IDENTICAL check per member; before,
/// the tuple form had no runtime path at all and leaked a raw `0`/`1`.
fn py_isinstance_runtime_check(value: &Expression, type_name: &str) -> Option<Expression> {
    let typeof_check = |name: &str| {
        Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(value.clone())))),
            right: Box::new(Expression::string(name)) })
    };
    let ref_test = |name: &str| {
        Expression::new(ExprKind::Binary {
            op: BinOp::InstanceOf,
            left: Box::new(value.clone()),
            right: Box::new(Expression::new(ExprKind::Ident(name.into()))) })
    };
    // `ref.test` pushes a raw wasm i32 — materialize a real Python bool.
    let as_bool = |e: Expression| {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(e),
            then: Box::new(Expression::bool(true)),
            else_: Box::new(Expression::bool(false)) })
    };
    let member = |field: &str| {
        Expression::new(ExprKind::Index {
            object: Box::new(value.clone()),
            index: Box::new(Expression::string(field)),
            null_safe: false })
    };
    let and = |l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(l),
            right: Box::new(r) })
    };
    let or = |l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(l),
            right: Box::new(r) })
    };
    let exception_type_check = |target: &str| {
        let exc_type = Expression::new(ExprKind::Index {
            object: Box::new(value.clone()),
            index: Box::new(Expression::string("__exception_type")),
            null_safe: false });
        let mut acc: Option<Expression> = None;
        for candidate in py_builtin_exception_names() {
            if py_builtin_subclass(candidate, target) == Some(true) {
                let one = Expression::new(ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(exc_type.clone()),
                    right: Box::new(Expression::string(candidate)) });
                acc = Some(match acc {
                    Some(prev) => or(prev, one),
                    None => one });
            }
        }
        acc
    };

    // A dict is EITHER Map-backed (`dict_literals_as_map = true`, the default
    // for literals) OR a legacy struct carrying a `__keys` array. Probing only
    // `__keys` — as this did before dicts became Maps — reports False for every
    // ordinary `{...}`. Sets are excluded (they trap on index) and so are
    // strings (they index by character).
    let dict_check = || {
        let keys_probe = Expression::new(ExprKind::Binary {
            op: BinOp::StrictNotEq,
            left: Box::new(member("__keys")),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))) });
        let not_set = Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(ref_test("Set")) });
        let struct_dict = and(
            and(typeof_check("object"), not_set),
            keys_probe,
        );
        Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(ref_test("Map")),
            right: Box::new(struct_dict) })
    };

    if let Some(check) = exception_type_check(type_name) {
        return Some(check);
    }

    Some(match type_name {
        "str" => typeof_check("string"),
        "FunctionType" | "LambdaType" => typeof_check("function"),
        "GeneratorType" | "CoroutineType" => Expression::bool(true),
        "bool" => typeof_check("boolean"),
        "float" => typeof_check("number"),
        // int includes bool — Python's bool IS an int subtype.
        "int" => Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(typeof_check("number")),
            right: Box::new(typeof_check("boolean")) }),
        // Both are ObjectKind::Array (the abstract WASM GC heap type), but a
        // tuple carries the `__tuple` tag (`tuple_literals_tagged`), which is
        // what repr/type() already key on. Without it `isinstance([1], tuple)`
        // is True and every list reprs as a tuple.
        "tuple" => as_bool(and(ref_test("array"), member("__tuple"))),
        "list" => as_bool(and(
            ref_test("array"),
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(member("__tuple")) }),
        )),
        "dict" => as_bool(dict_check()),
        "Sized" | "Iterable" => as_bool(or(
            or(ref_test("array"), typeof_check("string")),
            or(ref_test("Map"), ref_test("Set")),
        )),
        "Mapping" | "MutableMapping" => as_bool(dict_check()),
        "Sequence" => as_bool(or(ref_test("array"), typeof_check("string"))),
        "MutableSequence" => as_bool(and(
            ref_test("array"),
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(member("__tuple")) }),
        )),
        "Callable" => typeof_check("function"),
        "Iterator" => {
            if let ExprKind::Ident(name) = &value.kind
                && (is_iterator_var(name) || is_generator_var(name))
            {
                Expression::bool(true)
            } else {
                as_bool(Expression::new(ExprKind::Binary {
                    op: BinOp::StrictNotEq,
                    left: Box::new(member("next")),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))) }))
            }
        }
        "Generator" => {
            if let ExprKind::Ident(name) = &value.kind
                && is_generator_var(name)
            {
                Expression::bool(true)
            } else {
                as_bool(Expression::new(ExprKind::Binary {
                    op: BinOp::StrictNotEq,
                    left: Box::new(member("next")),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))) }))
            }
        }
        // A frozenset is a Set carrying the `__frozenset` tag (the same tag
        // `repr` reads to print `frozenset({…})`), so a bare `instanceof
        // frozenset` never matches.
        "frozenset" => as_bool(and(ref_test("Set"), member("__frozenset"))),
        "set" => as_bool(ref_test("Set")),
        _ => return None })
}

/// `pprint.<name>` → the injected prelude global (see [PPRINT_PRELUDE]).
fn pprint_module_member(field: &str) -> Option<&'static str> {
    Some(match field {
        "pformat" => "__pprint_pformat",
        "pprint" => "__pprint_pprint",
        "pp" => "__pprint_pp",
        "saferepr" => "__pprint_saferepr",
        "isreadable" => "__pprint_isreadable",
        "isrecursive" => "__pprint_isrecursive",
        "PrettyPrinter" => "__pprint_PrettyPrinter",
        _ => return None })
}

/// `shlex.<name>` → the injected prelude global (see [SHLEX_PRELUDE]).
fn shlex_module_member(field: &str) -> Option<&'static str> {
    Some(match field {
        "split" => "__py_shlex_split",
        "quote" => "__py_shlex_quote",
        "join" => "__py_shlex_join",
        "shlex" => "__py_shlex_class",
        _ => return None })
}

/// `textwrap.<name>` → the injected prelude global (see [TEXTWRAP_PRELUDE]).
fn textwrap_module_member(field: &str) -> Option<&'static str> {
    Some(match field {
        "wrap" => "__py_textwrap_wrap",
        "fill" => "__py_textwrap_fill",
        "dedent" => "__py_textwrap_dedent",
        "indent" => "__py_textwrap_indent",
        "shorten" => "__py_textwrap_shorten",
        "TextWrapper" => "__py_TextWrapper",
        _ => return None })
}

fn normalize_shlex_call_args(field: &str, mut args: Vec<Argument>) -> Vec<Argument> {
    if field == "shlex"
        && let Some(first) = args.first_mut()
        && let ExprKind::Ident(name) = &first.value.kind
        && let Some(text) = stringio_initial(name)
    {
        first.value = Expression::string(&text);
    }
    args
}

fn textwrap_default_arg(name: &str) -> Expression {
    match name {
        "text" => Expression::string(""),
        "width" => Expression::int(70),
        "initial_indent" | "subsequent_indent" => Expression::string(""),
        "break_long_words" | "break_on_hyphens" | "expand_tabs" | "replace_whitespace"
        | "drop_whitespace" => Expression::bool(true),
        "max_lines" => Expression::null(),
        "placeholder" => Expression::string(" [...]"),
        "prefix" => Expression::string(""),
        "predicate" => Expression::null(),
        _ => Expression::null() }
}

fn flatten_textwrap_args(field: &str, args: Vec<Argument>) -> Vec<Argument> {
    let params: &[&str] = match field {
        "wrap" | "fill" => &[
            "text",
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
        "TextWrapper" => &[
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
        "indent" => &["text", "prefix", "predicate"],
        "shorten" => &["text", "width", "placeholder"],
        _ => return args };

    if !args.iter().any(|a| a.name.is_some()) {
        return args;
    }

    let mut out: Vec<Argument> = params
        .iter()
        .map(|name| Argument::positional(textwrap_default_arg(name)))
        .collect();
    let mut pos = 0usize;
    for arg in args {
        if let Some(name) = &arg.name {
            if let Some(index) = params.iter().position(|p| *p == name) {
                out[index] = Argument::positional(arg.value);
            }
            // `tabsize` is accepted by TextWrapper; this prelude's
            // expandtabs path uses the default width, which is enough for the
            // current runtime surface.
        } else if pos < out.len() {
            out[pos] = Argument::positional(arg.value);
            pos += 1;
        }
    }
    while out
        .last()
        .is_some_and(|arg| matches!(arg.value.kind, ExprKind::Lit(Literal::Null)))
    {
        out.pop();
    }
    out
}

fn fold_textwrap_call(field: &str, args: &[Argument]) -> Option<Expression> {
    let lit = |index: usize| args.get(index).and_then(|a| resolve_string_const(&a.value));
    match field {
        "fill" => {
            let text = lit(0)?;
            let width = args.get(1).and_then(|a| expr_int(&a.value)).unwrap_or(70);
            let drop_whitespace = args.get(8).and_then(|a| expr_bool(&a.value)).unwrap_or(true);
            if !drop_whitespace && width >= text.len() as i64 {
                return Some(Expression::string(&text));
            }
            None
        }
        "dedent" => {
            let text = lit(0)?;
            let mut indent: Option<usize> = None;
            for line in text.split('\n') {
                if line.trim().is_empty() {
                    continue;
                }
                let n = line.chars().take_while(|c| *c == ' ' || *c == '\t').count();
                indent = Some(indent.map_or(n, |old| old.min(n)));
            }
            let Some(indent) = indent else {
                return Some(Expression::string(&text));
            };
            if indent == 0 {
                return Some(Expression::string(&text));
            }
            let out = text
                .split('\n')
                .map(|line| line.chars().skip(indent).collect::<String>())
                .collect::<Vec<_>>()
                .join("\n");
            Some(Expression::string(&out))
        }
        "indent" => {
            let text = lit(0)?;
            if text.is_empty() {
                return Some(Expression::string(""));
            }
            let prefix = lit(1)?;
            let predicate = args
                .get(2)
                .is_some_and(|a| !matches!(a.value.kind, ExprKind::Lit(Literal::Null)));
            let out = text
                .split('\n')
                .map(|line| {
                    if !predicate || !line.trim().is_empty() {
                        format!("{prefix}{line}")
                    } else {
                        line.to_string()
                    }
                })
                .collect::<Vec<_>>()
                .join("\n");
            Some(Expression::string(&out))
        }
        _ => None }
}

fn expr_int(e: &Expression) -> Option<i64> {
    match &e.kind {
        ExprKind::Lit(Literal::Int(n)) => Some(*n),
        _ => None }
}

fn expr_bool(e: &Expression) -> Option<bool> {
    match &e.kind {
        ExprKind::Lit(Literal::Bool(v)) => Some(*v),
        _ => None }
}

fn expr_str(e: &Expression) -> Option<String> {
    resolve_string_const(e)
}

fn rust_textwrap_wrap(
    text: &str,
    width: usize,
    initial_indent: &str,
    subsequent_indent: &str,
    break_long_words: bool,
    max_lines: Option<usize>,
    placeholder: &str,
) -> Vec<String> {
    let expanded = text.replace('\t', "    ");
    let words: Vec<&str> = expanded.split_whitespace().collect();
    if words.is_empty() {
        return Vec::new();
    }
    let mut lines = Vec::new();
    let mut cur = initial_indent.to_string();
    for word in words {
        if break_long_words && word.chars().count() > width {
            if !cur.trim().is_empty() {
                lines.push(cur.trim_end().to_string());
                cur = subsequent_indent.to_string();
            }
            let chars: Vec<char> = word.chars().collect();
            let mut i = 0usize;
            while i < chars.len() {
                let end = (i + width).min(chars.len());
                lines.push(chars[i..end].iter().collect());
                i = end;
            }
            continue;
        }
        let sep = if cur == initial_indent || cur == subsequent_indent {
            ""
        } else {
            " "
        };
        if cur.len() + sep.len() + word.len() <= width {
            cur.push_str(sep);
            cur.push_str(word);
        } else {
            if !cur.trim().is_empty() {
                lines.push(cur.trim_end().to_string());
            }
            cur = format!("{subsequent_indent}{word}");
        }
    }
    if !cur.trim().is_empty() {
        lines.push(cur.trim_end().to_string());
    }
    if let Some(max) = max_lines
        && lines.len() > max
    {
        lines.truncate(max);
        if let Some(last) = lines.last_mut() {
            let keep = width.saturating_sub(placeholder.len());
            *last = format!("{}{}", last.chars().take(keep).collect::<String>().trim_end(), placeholder);
        }
    }
    lines
}

fn fold_textwrapper_method(field: &str, settings: &[Expression], args: &[Argument]) -> Option<Expression> {
    let text = args.first().and_then(|a| expr_str(&a.value))?;
    let width = settings.first().and_then(expr_int).unwrap_or(70).max(1) as usize;
    let initial = settings.get(1).and_then(expr_str).unwrap_or_default();
    let subsequent = settings.get(2).and_then(expr_str).unwrap_or_default();
    let break_long = settings.get(3).and_then(expr_bool).unwrap_or(true);
    let max_lines = settings
        .get(8)
        .and_then(expr_int)
        .and_then(|n| usize::try_from(n).ok());
    let placeholder = settings
        .get(9)
        .and_then(expr_str)
        .unwrap_or_else(|| " [...]".to_string());
    let lines = rust_textwrap_wrap(
        &text,
        width,
        &initial,
        &subsequent,
        break_long,
        max_lines,
        &placeholder,
    );
    if field == "fill" {
        return Some(Expression::string(&lines.join("\n")));
    }
    Some(Expression::new(ExprKind::Array(
        lines
            .into_iter()
            .map(|line| ArrayElement {
                key: None,
                value: Expression::string(&line),
                spread: false,
                by_ref: false })
            .collect(),
    )))
}

fn rust_fnmatch_class_match(ch: char, chars: &[char], mut i: usize) -> Option<(bool, usize)> {
    let mut negate = false;
    if i < chars.len() && matches!(chars[i], '!' | '^') {
        negate = true;
        i += 1;
    }
    let mut matched = false;
    let mut closed = false;
    while i < chars.len() {
        if chars[i] == ']' {
            closed = true;
            i += 1;
            break;
        }
        if i + 2 < chars.len() && chars[i + 1] == '-' && chars[i + 2] != ']' {
            if chars[i] <= ch && ch <= chars[i + 2] {
                matched = true;
            }
            i += 3;
        } else {
            if chars[i] == ch {
                matched = true;
            }
            i += 1;
        }
    }
    closed.then_some((if negate { !matched } else { matched }, i))
}

fn rust_fnmatch(name: &str, pattern: &str, case_sensitive: bool) -> bool {
    let name = if case_sensitive {
        name.to_string()
    } else {
        name.to_lowercase()
    };
    let pattern = if case_sensitive {
        pattern.to_string()
    } else {
        pattern.to_lowercase()
    };
    let n: Vec<char> = name.chars().collect();
    let p: Vec<char> = pattern.chars().collect();
    let mut memo = std::collections::HashMap::<(usize, usize), bool>::new();
    fn rec(
        n: &[char],
        p: &[char],
        ni: usize,
        pi: usize,
        memo: &mut std::collections::HashMap<(usize, usize), bool>,
    ) -> bool {
        if let Some(v) = memo.get(&(ni, pi)) {
            return *v;
        }
        let out = if pi == p.len() {
            ni == n.len()
        } else if p[pi] == '*' {
            rec(n, p, ni, pi + 1, memo) || (ni < n.len() && rec(n, p, ni + 1, pi, memo))
        } else if ni < n.len() && p[pi] == '?' {
            rec(n, p, ni + 1, pi + 1, memo)
        } else if ni < n.len() && p[pi] == '[' {
            if let Some((ok, next_pi)) = rust_fnmatch_class_match(n[ni], p, pi + 1) {
                ok && rec(n, p, ni + 1, next_pi, memo)
            } else {
                p[pi] == n[ni] && rec(n, p, ni + 1, pi + 1, memo)
            }
        } else {
            ni < n.len() && p[pi] == n[ni] && rec(n, p, ni + 1, pi + 1, memo)
        };
        memo.insert((ni, pi), out);
        out
    }
    rec(&n, &p, 0, 0, &mut memo)
}

/// CPython `fnmatch.translate` — note the DIALECT: `(?s:BODY)\\Z`, not
/// `^BODY$`. Measured, `fnmatch.translate("*.py")` is `(?s:.*\\.py)\\Z`; this
/// emitted the ECMA-shaped anchors instead, so python's own regex idiom was
/// wrong wherever the result was used or printed.
///
/// [`fnmatch_generated_regex_to_pattern`] is the inverse and must stay in step.
fn rust_fnmatch_translate(pattern: &str) -> String {
    let mut out = String::from("(?s:");
    let chars: Vec<char> = pattern.chars().collect();
    let mut i = 0usize;
    while i < chars.len() {
        match chars[i] {
            '*' => out.push_str(".*"),
            '?' => out.push('.'),
            '[' => {
                let start = i;
                i += 1;
                if i < chars.len() && matches!(chars[i], '!' | '^') {
                    out.push_str("[^");
                    i += 1;
                } else {
                    out.push('[');
                }
                let mut closed = false;
                while i < chars.len() {
                    if chars[i] == ']' {
                        closed = true;
                        break;
                    }
                    if chars[i] == '\\' {
                        out.push('\\');
                    }
                    out.push(chars[i]);
                    i += 1;
                }
                if closed {
                    out.push(']');
                } else {
                    out.push_str("\\[");
                    i = start;
                }
            }
            c if matches!(c, '.' | '\\' | '+' | '(' | ')' | '|' | '^' | '$' | '{' | '}') => {
                out.push('\\');
                out.push(c);
            }
            c => out.push(c) }
        i += 1;
    }
    out.push_str(")\\Z");
    out
}

fn fold_fnmatch_call(field: &str, args: &[Argument]) -> Option<Expression> {
    match field {
        "fnmatch" | "fnmatchcase" if args.len() >= 2 => {
            let name = resolve_string_const(&args[0].value)?;
            let pat = resolve_string_const(&args[1].value)?;
            Some(Expression::bool(rust_fnmatch(
                &name,
                &pat,
                field == "fnmatchcase",
            )))
        }
        "filter" if args.len() >= 2 => {
            let names = resolve_string_array_const(&args[0].value)?;
            let pat = resolve_string_const(&args[1].value)?;
            let mut out = Vec::new();
            for name in names {
                if rust_fnmatch(&name, &pat, false) {
                    out.push(ArrayElement {
                        key: None,
                        value: Expression::string(&name),
                        spread: false,
                        by_ref: false });
                }
            }
            Some(Expression::new(ExprKind::Array(out)))
        }
        "translate" if args.len() == 1 => {
            let pat = resolve_string_const(&args[0].value)?;
            Some(Expression::string(&rust_fnmatch_translate(&pat)))
        }
        _ => None }
}

fn fnmatch_generated_regex_to_pattern(regex: &str) -> Option<String> {
    // The inverse of `rust_fnmatch_translate`, so it reads python's dialect.
    let body = regex.strip_prefix("(?s:")?.strip_suffix(")\\Z")?;
    let chars: Vec<char> = body.chars().collect();
    let mut out = String::new();
    let mut i = 0usize;
    while i < chars.len() {
        if i + 1 < chars.len() && chars[i] == '.' && chars[i + 1] == '*' {
            out.push('*');
            i += 2;
            continue;
        }
        match chars[i] {
            '.' => out.push('?'),
            '\\' if i + 1 < chars.len() => {
                out.push(chars[i + 1]);
                i += 2;
                continue;
            }
            '[' => {
                out.push('[');
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    if chars[i] == ']' {
                        break;
                    }
                    i += 1;
                }
            }
            c => out.push(c) }
        i += 1;
    }
    Some(out)
}

fn fold_re_call(field: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(field, "match" | "search") || args.len() < 2 {
        return None;
    }
    let regex = resolve_string_const(&args[0].value)?;
    let text = resolve_string_const(&args[1].value)?;
    let pat = fnmatch_generated_regex_to_pattern(&regex)?;
    Some(Expression::bool(rust_fnmatch(&text, &pat, true)))
}

/// DB-API 2.0 module constants for `sqlite3` (static mount → compile-time).
fn sqlite3_module_constant(field: &str) -> Option<Literal> {
    Some(match field {
        "paramstyle" => Literal::Str("qmark".into()),
        "apilevel" => Literal::Str("2.0".into()),
        "threadsafety" => Literal::Int(1),
        "version" => Literal::Str("2.6.0".into()),
        "sqlite_version" => Literal::Str("3.40.0".into()),
        _ => return None })
}

/// Map a sqlite cursor/connection method name to its `__sql_*` builtin.
fn sql_method_builtin(field: &str) -> Option<&'static str> {
    Some(match field {
        "cursor" => "__sql_cursor",
        "execute" => "__sql_execute",
        "executemany" => "__sql_executemany",
        "fetchall" => "__sql_fetchall",
        "fetchone" => "__sql_fetchone",
        "commit" => "__sql_commit",
        "rollback" => "__sql_rollback",
        "close" => "__sql_close",
        _ => return None })
}

/// True when `e` is a sqlite connection/cursor handle: a tracked variable, or a
/// call to a `__sql_*` builtin that returns a handle (so `conn.execute(...)
/// .fetchone()` and `sqlite3.connect(...).execute(...)` chain).
fn is_sql_handle_expr(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Ident(name) => is_sql_var(name),
        ExprKind::Call { callee, .. } => matches!(&callee.kind, ExprKind::Ident(f)
            if matches!(
                f.as_str(),
                "__sql_connect" | "__sql_cursor" | "__sql_execute" | "__sql_executemany"
            )),
        _ => false }
}

/// `sqlite3.connect(...)` → `__sql_connect(...)`, and `<handle>.method(...)` →
/// `__sql_method(<handle>, ...)`. `args` are already desugared. Returns `None`
/// when the receiver is not a sqlite handle, so unrelated `.close()`/`.execute()`
/// fall through to normal method dispatch.
fn rewrite_sqlite_call(
    object: &Expression,
    field: &str,
    args: Vec<Argument>,
    optional: bool,
) -> Option<Expression> {
    // `sqlite3.connect(path)`
    if let ExprKind::Ident(module) = &object.kind {
        if module == "sqlite3" {
            if field == "connect" && !args.is_empty() {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("__sql_connect")),
                    args,
                    optional }));
            }
            // `sqlite3.Binary(b)` is identity on the bytes it wraps.
            if field == "Binary" && args.len() == 1 {
                return Some(desugar_member_reads(args.into_iter().next().unwrap().value));
            }
        }
    }
    let builtin = sql_method_builtin(field)?;
    // Desugar the receiver first so a chained producer call (`conn.execute(...)`)
    // becomes a `__sql_*` call we can recognize as a handle.
    let recv = desugar_member_reads(object.clone());
    if !is_sql_handle_expr(&recv) {
        return None;
    }
    let mut call_args = Vec::with_capacity(args.len() + 1);
    call_args.push(Argument::positional(recv));
    call_args.extend(args);
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(builtin)),
        args: call_args,
        optional }))
}

fn note_from_imported_module(name: &str) {
    PY_FROM_IMPORTED_MODULES.with(|m| m.borrow_mut().insert(name.to_string()));
}

fn is_from_imported_module(name: &str) -> bool {
    PY_FROM_IMPORTED_MODULES.with(|m| m.borrow().contains(name))
}

fn note_float_returning_import(module: &str, imported: &str, local: &str) {
    let is_float = match module {
        "math" => FLOAT_MATH_FNS.contains(&imported),
        "statistics" => FLOAT_STATISTICS_FNS.contains(&imported),
        _ => false };
    if is_float {
        PY_FLOAT_RETURNING_IMPORTS.with(|m| {
            m.borrow_mut().insert(local.to_string());
        });
    }
}

fn is_float_returning_import(name: &str) -> bool {
    PY_FLOAT_RETURNING_IMPORTS.with(|m| m.borrow().contains(name))
}

fn note_imported_module(name: &str) {
    // Track both the full dotted path's first segment and any alias so that a
    // bare `mod.CONST` read is left as namespace access, not turned into a
    // subscript.
    let first = name.split('.').next().unwrap_or(name).trim();
    if !first.is_empty() {
        PY_IMPORTED_MODULES.with(|m| m.borrow_mut().insert(first.to_string()));
    }
}

fn is_imported_module(name: &str) -> bool {
    PY_IMPORTED_MODULES.with(|m| m.borrow().contains(name))
}

thread_local! {
    /// `m = importlib.import_module('json')` / `m = json` — locals aliased
    /// to a mounted module. Member access substitutes the module root so
    /// `m.dumps(...)` compiles exactly like `json.dumps(...)`.
    static PY_MODULE_ALIASES: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
}

/// True when `receiver.field` names a class defined by an injected stdlib
/// prelude (e.g. `io.StringIO`, `configparser.ConfigParser`). Such a
/// `module.Class(...)` must construct the bare global class rather than compile
/// to a method call that passes the module object as the first argument.
fn prelude_module_class(receiver: &ExprKind, field: &str) -> Option<String> {
    let ExprKind::Ident(module) = receiver else {
        return None;
    };
    match (module.as_str(), field) {
        // `socket.socket` is the one case where the class and its module share
        // a name, so the global cannot be called `socket` — the walker already
        // tracks that identifier as an imported module. Map it to a distinct
        // global instead of reusing the field name.
        ("socket", "socket") => Some("VybeSocketImpl".to_string()),
        ("io", "StringIO")
        | ("io", "BytesIO")
        | ("configparser", "ConfigParser")
        | ("configparser", "RawConfigParser") => Some(field.to_string()),
        _ => None }
}

fn note_module_alias(alias: &str, module: &str) {
    PY_MODULE_ALIASES.with(|m| {
        m.borrow_mut().insert(alias.to_string(), module.to_string());
    });
    note_imported_module(alias);
}

fn resolve_module_alias(name: &str) -> Option<String> {
    PY_MODULE_ALIASES.with(|m| m.borrow().get(name).cloned())
}

fn object_string_property(e: &Expression, key: &str) -> Option<String> {
    let ExprKind::Object(props) = &e.kind else {
        return None;
    };
    props.iter().find_map(|prop| {
        let ObjectProperty::KeyValue { key: k, value } = prop else {
            return None;
        };
        let ExprKind::Lit(Literal::Str(s)) = &k.kind else {
            return None;
        };
        if s.to_string() != key {
            return None;
        }
        match &value.kind {
            ExprKind::Lit(Literal::Str(s)) => Some(s.to_string()),
            _ => None }
    })
}

fn note_dynamic_module_var(var: &str, module: &str) {
    PY_DYNAMIC_MODULE_VARS.with(|m| {
        m.borrow_mut().insert(var.to_string(), module.to_string());
    });
    note_imported_module(module);
}

fn dynamic_module_for_var(var: &str) -> Option<String> {
    PY_DYNAMIC_MODULE_VARS.with(|m| m.borrow().get(var).cloned())
}

fn note_dynamic_module_registry(module: &str, var: &str) {
    PY_DYNAMIC_MODULE_REGISTRY.with(|m| {
        m.borrow_mut().insert(module.to_string(), var.to_string());
    });
    note_dynamic_module_var(var, module);
}

fn dynamic_module_registry_var(module: &str) -> Option<String> {
    PY_DYNAMIC_MODULE_REGISTRY.with(|m| m.borrow().get(module).cloned())
}

fn dynamic_module_attr_target(target: &Expression) -> Option<(String, String)> {
    match &target.kind {
        ExprKind::Member { object, field, .. } => {
            if let ExprKind::Ident(var) = &object.kind {
                dynamic_module_for_var(var).map(|module| (module, field.clone()))
            } else {
                None
            }
        }
        ExprKind::Index { object, index, .. } => {
            let ExprKind::Ident(var) = &object.kind else {
                return None;
            };
            let ExprKind::Lit(Literal::Str(attr)) = &index.kind else {
                return None;
            };
            dynamic_module_for_var(var).map(|module| (module, attr.to_string()))
        }
        _ => None }
}

fn literal_string_array(value: &Expression) -> Option<Vec<String>> {
    let ExprKind::Array(elems) = &value.kind else {
        return None;
    };
    let mut out = Vec::with_capacity(elems.len());
    for elem in elems {
        let ExprKind::Lit(Literal::Str(s)) = &elem.value.kind else {
            return None;
        };
        out.push(s.to_string());
    }
    Some(out)
}

fn note_dynamic_module_attr(module: &str, attr: &str, value: Expression) {
    if attr == "__all__" {
        if let Some(names) = literal_string_array(&value) {
            PY_DYNAMIC_MODULE_ALL.with(|m| {
                m.borrow_mut().insert(module.to_string(), names);
            });
        }
    }
    PY_DYNAMIC_MODULE_ATTRS.with(|m| {
        let mut map = m.borrow_mut();
        let attrs = map.entry(module.to_string()).or_default();
        if let Some((_, existing)) = attrs.iter_mut().find(|(name, _)| name == attr) {
            *existing = value;
        } else {
            attrs.push((attr.to_string(), value));
        }
    });
}

fn dynamic_module_attr(module: &str, attr: &str) -> Option<Expression> {
    PY_DYNAMIC_MODULE_ATTRS.with(|m| {
        m.borrow()
            .get(module)
            .and_then(|attrs| attrs.iter().find(|(name, _)| name == attr).map(|(_, v)| v.clone()))
    })
}

fn dynamic_module_all(module: &str) -> Option<Vec<String>> {
    PY_DYNAMIC_MODULE_ALL.with(|m| m.borrow().get(module).cloned())
}

fn py_module_metadata_attr(module_name: &str, field: &str) -> Option<Expression> {
    let string = |s: String| Expression::new(ExprKind::Lit(Literal::Str(s.into())));
    Some(match field {
        "__name__" => string(module_name.to_string()),
        "__file__" => string(format!("<{module_name}>")),
        "__doc__" => string(String::new()),
        "__package__" => {
            let package = module_name
                .rsplit_once('.')
                .map(|(pkg, _)| pkg)
                .unwrap_or("");
            string(package.to_string())
        }
        "__loader__" => Expression::new(ExprKind::Object(vec![])),
        "__spec__" => py_module_spec_object(module_name),
        _ => return None })
}

fn py_module_spec_object(module_name: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str("name".into()))),
            value: Expression::new(ExprKind::Lit(Literal::Str(module_name.to_string().into()))) },
        ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str("loader".into()))),
            value: Expression::new(ExprKind::Object(vec![])) },
    ]))
}

fn dynamic_module_import_stmts(module: &str, local: &str) -> Option<Vec<Statement>> {
    let source = dynamic_module_registry_var(module)?;
    let mut stmts = Vec::new();
    if local != source {
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Ident(local.to_string()))],
            value: Expression::new(ExprKind::Ident(source.clone())), by_ref: false }));
    }
    note_dynamic_module_var(local, module);
    Some(stmts)
}

fn dynamic_module_star_import_stmts(module: &str) -> Option<Vec<Statement>> {
    let names = dynamic_module_all(module)?;
    let mut stmts = Vec::new();
    for name in names {
        if name.starts_with('_') {
            continue;
        }
        let value = dynamic_module_attr(module, &name)?;
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Ident(name))],
            value, by_ref: false }));
    }
    Some(stmts)
}

fn note_string_const(name: &str, value: &str) {
    PY_STRING_CONSTS.with(|m| {
        m.borrow_mut().insert(name.to_string(), value.to_string());
    });
}

fn clear_string_const(name: &str) {
    PY_STRING_CONSTS.with(|m| {
        m.borrow_mut().remove(name);
    });
}

fn note_none_var(name: &str) {
    PY_NONE_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn clear_none_var(name: &str) {
    PY_NONE_VARS.with(|m| {
        m.borrow_mut().remove(name);
    });
}

fn expr_is_tracked_none(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Lit(Literal::Null) => true,
        ExprKind::Ident(name) => PY_NONE_VARS.with(|m| m.borrow().contains(name)),
        _ => false }
}

fn resolve_string_const(e: &Expression) -> Option<String> {
    match &e.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.to_string()),
        ExprKind::Ident(name) => PY_STRING_CONSTS.with(|m| m.borrow().get(name).cloned()),
        _ => None }
}

fn mapping_proxy_source(name: &str) -> Option<Expression> {
    PY_MAPPING_PROXY_VARS.with(|m| m.borrow().get(name).cloned())
}

fn note_mapping_proxy_var(name: &str, source: Expression) {
    PY_MAPPING_PROXY_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string(), source);
    });
}

fn mapping_proxy_ctor_arg(value: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &value.kind else {
        return None;
    };
    let is_ctor = match &callee.kind {
        ExprKind::Ident(n) => n == "MappingProxyType",
        ExprKind::Member { object, field, .. } => {
            matches!(&object.kind, ExprKind::Ident(n) if n == "types")
                && field == "MappingProxyType"
        }
        _ => false };
    if is_ctor && args.len() == 1 {
        Some(args[0].value.clone())
    } else {
        None
    }
}

fn note_simple_namespace_var(name: &str) {
    PY_SIMPLE_NAMESPACE_VARS.with(|m| {
        m.borrow_mut().insert(name.to_string());
    });
}

fn is_simple_namespace_var(name: &str) -> bool {
    PY_SIMPLE_NAMESPACE_VARS.with(|m| m.borrow().contains(name))
}

fn simple_namespace_ctor_object(value: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &value.kind else {
        return None;
    };
    let is_ctor = match &callee.kind {
        ExprKind::Ident(n) => n == "SimpleNamespace",
        ExprKind::Member { object, field, .. } => {
            matches!(&object.kind, ExprKind::Ident(n) if n == "types")
                && field == "SimpleNamespace"
        }
        _ => false };
    if !is_ctor {
        return None;
    }
    let mut props = Vec::new();
    for arg in args {
        if arg.spread {
            props.push(ObjectProperty::Spread(arg.value.clone()));
        } else if let Some(name) = &arg.name {
            props.push(ObjectProperty::KeyValue {
                key: Expression::string(name),
                value: arg.value.clone() });
        }
    }
    Some(Expression::new(ExprKind::Object(props)))
}

fn keyword_object(args: &[Argument]) -> Option<Expression> {
    let props: Vec<ObjectProperty> = args
        .iter()
        .filter_map(|arg| {
            let name = arg.name.as_ref()?;
            Some(ObjectProperty::KeyValue {
                key: Expression::string(name),
                value: desugar_member_reads(arg.value.clone()) })
        })
        .collect();
    (!props.is_empty()).then(|| Expression::new(ExprKind::Object(props)))
}

fn collections_ctor_call(name: &str, args: &[Argument]) -> Option<Expression> {
    let positional: Vec<Expression> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| desugar_member_reads(a.value.clone()))
        .collect();
    match name {
        "Counter" | "__py_counter_new" => {
            let iterable = positional
                .first()
                .cloned()
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            let kws = keyword_object(args)
                .or_else(|| {
                    if name == "__py_counter_new" {
                        positional.get(1).cloned()
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            Some(call_ident("__py_counter_new", vec![iterable, kws]))
        }
        "defaultdict" | "__py_defaultdict" => {
            let factory = positional
                .first()
                .cloned()
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            let initial = positional
                .get(1)
                .cloned()
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            Some(call_ident("__py_defaultdict", vec![factory, initial]))
        }
        "deque" | "__py_deque" => {
            let iterable = positional
                .first()
                .cloned()
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            let maxlen = args
                .iter()
                .find(|a| a.name.as_deref() == Some("maxlen"))
                .map(|a| desugar_member_reads(a.value.clone()))
                .or_else(|| positional.get(1).cloned())
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            Some(call_ident("__py_deque", vec![iterable, maxlen]))
        }
        "ChainMap" | "__py_chainmap_new" => {
            let call_args = positional.into_iter().map(Argument::positional).collect();
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__py_chainmap_new")),
                args: call_args,
                optional: false }))
        }
        "UserDict" | "__py_userdict" => Some(
            positional
                .first()
                .cloned()
                .unwrap_or_else(|| Expression::new(ExprKind::Object(Vec::new()))),
        ),
        "UserList" | "__py_userlist" => Some(call_ident(
            "__py_userlist",
            vec![positional
                .first()
                .cloned()
                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))],
        )),
        "UserString" | "__py_userstring" => Some(call_ident(
            "str",
            vec![positional
                .first()
                .cloned()
                .unwrap_or_else(|| Expression::string(""))],
        )),
        _ => None }
}

fn py_module_callable_member(module: &str, attr: &str) -> Option<Expression> {
    let max_args = match (module, attr) {
        ("collections", "Counter" | "deque" | "OrderedDict") => 1,
        ("json", "dumps" | "loads" | "dump" | "load") => 1,
        _ => return None };
    let callable = Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(module)),
        field: attr.into(),
        null_safe: false });
    let params: Vec<Param> = (0..max_args)
        .map(|i| Param {
            name: format!("__arg{i}"),
            type_hint: None,
            default: Some(Expression::new(ExprKind::Lit(Literal::Null))),
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: true,
            is_nullable: true })
        .collect();
    let args: Vec<Argument> = (0..max_args)
        .map(|i| {
            let name = format!("__arg{i}");
            Argument::positional(Expression::ident(&name))
        })
        .collect();
    Some(Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(callable),
            args,
            optional: false }))),
        is_async: false,
        captures: vec![] }))
}

/// The stdlib universe this implementation can mount. `import X` for an
/// X outside it raises ImportError at the import site (CPython behavior)
/// instead of silently binding nothing.
fn py_known_module(root: &str) -> bool {
    py_module_surface(root).is_some()
        || matches!(
            root,
            "math"
                | "cmath"
                | "os"
                | "sys"
                | "json"
                | "re"
                | "random"
                | "time"
                | "datetime"
                | "calendar"
                | "zoneinfo"
                | "collections"
                | "itertools"
                | "functools"
                | "operator"
                | "string"
                | "textwrap"
                | "unicodedata"
                | "struct"
                | "codecs"
                | "io"
                | "abc"
                | "numbers"
                | "decimal"
                | "fractions"
                | "statistics"
                | "array"
                | "bisect"
                | "heapq"
                | "copy"
                | "pprint"
                | "enum"
                | "typing"
                | "dataclasses"
                | "contextlib"
                | "traceback"
                | "warnings"
                | "gc"
                | "inspect"
                | "builtins"
                | "pickle"
                | "hashlib"
                | "hmac"
                | "secrets"
                | "uuid"
                | "base64"
                | "binascii"
                | "shutil"
                | "tempfile"
                | "pathlib"
                | "stat"
                | "subprocess"
                | "threading"
                | "queue"
                | "socket"
                | "socketserver"
                | "ipaddress"
                | "ssl"
                | "select"
                | "signal"
                | "errno"
                | "platform"
                | "getpass"
                | "logging"
                | "argparse"
                | "unittest"
                | "doctest"
                | "timeit"
                | "csv"
                | "configparser"
                | "sqlite3"
                | "zlib"
                | "gzip"
                | "zipfile"
                | "tarfile"
                | "asyncio"
                | "weakref"
                | "atexit"
                | "keyword"
                | "token"
                | "ast"
                | "dis"
                | "sysconfig"
                | "__future__"
                | "urllib"
                | "http"
                | "email"
                | "xml"
                | "html"
        )
}

/// `raise ImportError('<msg>')` at the import site.
fn py_import_error_stmt(msg: &str) -> Statement {
    let err = Expression::ident("__py_import_error");
    let mut stmts = vec![Statement::new(StmtKind::Assign {
        targets: vec![err.clone()],
        value: Expression::new(ExprKind::New {
            class: Box::new(Expression::new(ExprKind::Ident("ImportError".into()))),
            args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                Literal::Str(msg.into()),
            )))] }), by_ref: false })];
    if let Some(name) = msg
        .strip_prefix("No module named '")
        .and_then(|s| s.strip_suffix("'"))
    {
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Index {
                object: Box::new(err.clone()),
                index: Box::new(Expression::new(ExprKind::Lit(Literal::Str("name".into())))),
                null_safe: false })],
            value: Expression::new(ExprKind::Lit(Literal::Str(name.into()))), by_ref: false }));
    }
    stmts.push(Statement::new(StmtKind::Throw {
        expr: Some(err),
        cause: None }));
    Statement::new(StmtKind::Block(stmts))
}

/// Python-facing names of mounted host modules that differ from the
/// canonical host export names — normalized as plain AST assignments at
/// import (`json['dumps'] = json['stringify']`), so the surface exists on
/// the runtime namespace object for reflection (dir/getattr/values) with
/// ZERO primitives/runtime machinery. JS never needs this: its names ARE
/// the canonical names.
fn py_module_renames(module: &str) -> Option<&'static [(&'static str, &'static str)]> {
    Some(match module {
        "json" => &[
            ("dumps", "stringify"),
            ("loads", "parse"),
            ("dump", "stringify"),
            ("load", "parse"),
        ],
        _ => return None })
}

/// Statements normalizing a module's Python-facing surface onto its
/// runtime namespace object, guarded so installer-less harnesses skip.
fn py_module_rename_stmts(module: &str) -> Vec<Statement> {
    let Some(renames) = py_module_renames(module) else {
        return Vec::new();
    };
    let mut assigns = Vec::new();
    for (py_name, canonical) in renames {
        assigns.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Index {
                object: Box::new(Expression::new(ExprKind::Ident(module.to_string()))),
                index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                    (*py_name).into(),
                )))),
                null_safe: false })],
            value: Expression::new(ExprKind::Index {
                object: Box::new(Expression::new(ExprKind::Ident(module.to_string()))),
                index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                    (*canonical).into(),
                )))),
                null_safe: false }), by_ref: false }));
    }
    // if typeof(module) != "undefined": <assigns>
    vec![Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(
                Expression::new(ExprKind::Ident(module.to_string())),
            )))),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                "undefined".into(),
            )))) }),
        then_body: assigns,
        elifs: Vec::new(),
        else_body: None })]
}

/// Static surfaces of stdlib modules the walker mounts — `hasattr(mod,
/// 'lit')` resolves at compile time against this table (same category as
/// `find_spec`/`sys.modules`: the mounts are static, so is the answer).
fn py_module_surface(module: &str) -> Option<&'static [&'static str]> {
    Some(match module {
        "importlib" => &[
            "import_module",
            "reload",
            "invalidate_caches",
            "util",
            "machinery",
            "resources",
            "metadata",
            "abc",
        ],
        "importlib.util" => &["find_spec", "module_from_spec", "spec_from_loader"],
        "importlib.machinery" => &["SourceFileLoader", "ExtensionFileLoader", "ModuleSpec"],
        "importlib.resources" => &["files", "read_text", "read_binary"],
        "importlib.metadata" => &["version", "distributions", "metadata"],
        "importlib.abc" => &["MetaPathFinder", "Loader", "PathEntryFinder"],
        "runpy" => &["run_module", "run_path"],
        "encodings" => &["utf_8", "ascii", "latin_1"],
        "pkgutil" => &["iter_modules", "walk_packages", "get_data"],
        "zipimport" => &["zipimporter", "ZipImportError"],
        "types" => &[
            "ModuleType",
            "SimpleNamespace",
            "MappingProxyType",
            "MethodType",
            "FunctionType",
            "LambdaType",
            "GeneratorType",
            "CoroutineType",
            "DynamicClassAttribute",
            "new_class",
            "resolve_bases",
        ],
        "collections" => &[
            "Counter",
            "defaultdict",
            "deque",
            "OrderedDict",
            "ChainMap",
            "namedtuple",
            "UserDict",
            "UserList",
            "UserString",
            "abc",
        ],
        "collections.abc" => &[
            "Sized",
            "Mapping",
            "Sequence",
            "Iterable",
            "Iterator",
            "Generator",
            "Callable",
            "Set",
            "MutableMapping",
            "MutableSequence",
        ],
        "email" => &["mime"],
        "email.mime" => &["text"],
        "email.mime.text" => &["MIMEText"],
        "xml" => &["etree"],
        "xml.etree" => &["ElementTree"],
        "xml.etree.ElementTree" => &["Element", "SubElement", "fromstring", "tostring"],
        "json" => &["dumps", "loads", "dump", "load"],
        "functools" => &["wraps", "reduce"],
        "shlex" => &["split", "quote", "join", "shlex"],
        "textwrap" => &["wrap", "fill", "dedent", "indent", "shorten", "TextWrapper"],
        "zoneinfo" => &["ZoneInfo", "available_timezones", "ZoneInfoNotFoundError"],
        "glob" => &["glob", "iglob", "escape", "has_magic"],
        "fnmatch" => &["fnmatch", "fnmatchcase", "filter", "translate"],
        "keyword" => &["iskeyword", "issoftkeyword", "kwlist", "softkwlist"],
        "mimetypes" => &[
            "guess_type",
            "guess_extension",
            "guess_all_extensions",
            "add_type",
            "init",
            "MimeTypes",
            "types_map",
            "encodings_map",
            "suffix_map",
        ],
        "getopt" => &["getopt", "gnu_getopt", "GetoptError", "error"],
        "os.path" => &[
            // string-math functions (prelude helpers)
            "join",
            "split",
            "splitext",
            "splitroot",
            "splitdrive",
            "basename",
            "dirname",
            "normpath",
            "isabs",
            "relpath",
            "commonprefix",
            "commonpath",
            "normcase",
            "realpath",
            "abspath",
            "expanduser",
            "expandvars",
            // FS predicates (host-backed) + query stubs the tests probe
            "exists",
            "lexists",
            "isfile",
            "isdir",
            "islink",
            "ismount",
            "samefile",
            "sameopenfile",
            "getsize",
            "getmtime",
            "getatime",
            "getctime",
            "isblock",
            "ischar",
            "isfifo",
            "issocket",
            // constants
            "sep",
            "altsep",
            "pathsep",
            "extsep",
            "curdir",
            "pardir",
            "defpath",
            "devnull",
        ],
        "sqlite3" => &[
            "connect",
            "Connection",
            "Cursor",
            "Row",
            "Binary",
            "Error",
            "DatabaseError",
            "IntegrityError",
            "OperationalError",
            "ProgrammingError",
            "register_adapter",
            "register_converter",
            "complete_statement",
            "paramstyle",
            "apilevel",
            "threadsafety",
            "version",
            "sqlite_version",
            "PARSE_DECLTYPES",
            "PARSE_COLNAMES",
        ],
        _ => return None })
}

thread_local! {
    // Names of classes declared in the module. Populated as each `class` is
    // walked (see `walk_class_def`), used to normalise a bare construction call
    // `ClassName(...)` to `ExprKind::New` — the SAME shape JS (`new F()`) and
    // PHP (`new F()`) produce. Python's grammar writes construction without a
    // `new` keyword, so it otherwise parses as a plain `Call`; normalising it
    // here lets a constructed instance used as a receiver (`F().m(...)`)
    // dispatch through the identical instance path as every other language.
    static PY_DEFINED_CLASSES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_DEFINED_FUNCTIONS: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_CALLABLE_CLASSES: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_CLASSES_WITH_INIT: std::cell::RefCell<std::collections::HashSet<String>> =
        std::cell::RefCell::new(std::collections::HashSet::new());
    static PY_CLASS_PARENTS: std::cell::RefCell<std::collections::HashMap<String, Vec<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_CLASS_ATTRS: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashSet<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_CLASS_DATA_ATTRS: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashSet<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_INSTANCE_CLASSES: std::cell::RefCell<std::collections::HashMap<String, String>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_INSTANCE_ATTRS: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashSet<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_ASSIGN_TARGET_DEPTH: std::cell::Cell<usize> = std::cell::Cell::new(0);
}

fn in_assignment_target() -> bool {
    PY_ASSIGN_TARGET_DEPTH.with(|d| d.get() > 0)
}

fn with_assignment_target<T>(f: impl FnOnce() -> T) -> T {
    PY_ASSIGN_TARGET_DEPTH.with(|d| d.set(d.get() + 1));
    let out = f();
    PY_ASSIGN_TARGET_DEPTH.with(|d| d.set(d.get().saturating_sub(1)));
    out
}

fn note_defined_class(name: &str) {
    if !name.is_empty() {
        PY_DEFINED_CLASSES.with(|m| {
            m.borrow_mut().insert(name.to_string());
        });
    }
}

fn note_defined_function(name: &str) {
    if !name.is_empty() {
        PY_DEFINED_FUNCTIONS.with(|m| {
            m.borrow_mut().insert(name.to_string());
        });
    }
}

fn is_defined_function(name: &str) -> bool {
    PY_DEFINED_FUNCTIONS.with(|m| m.borrow().contains(name))
}

fn note_class_parents(name: &str, parents: &[String]) {
    if !name.is_empty() {
        PY_CLASS_PARENTS.with(|m| {
            m.borrow_mut().insert(name.to_string(), parents.to_vec());
        });
    }
}

fn note_callable_class(name: &str) {
    if !name.is_empty() {
        PY_CALLABLE_CLASSES.with(|m| {
            m.borrow_mut().insert(name.to_string());
        });
    }
}

fn is_callable_class(name: &str) -> bool {
    PY_CALLABLE_CLASSES.with(|m| m.borrow().contains(name))
}

fn note_class_with_init(name: &str) {
    if !name.is_empty() {
        PY_CLASSES_WITH_INIT.with(|m| {
            m.borrow_mut().insert(name.to_string());
        });
    }
}

fn class_has_init(name: &str) -> bool {
    PY_CLASSES_WITH_INIT.with(|m| m.borrow().contains(name))
}

fn class_parent_has_init(name: &str) -> bool {
    PY_CLASS_PARENTS.with(|m| {
        let parents = m.borrow();
        let mut stack = parents.get(name).cloned().unwrap_or_default();
        let mut seen = std::collections::HashSet::new();
        while let Some(parent) = stack.pop() {
            if !seen.insert(parent.clone()) {
                continue;
            }
            if class_has_init(&parent) {
                return true;
            }
            if let Some(more) = parents.get(&parent) {
                stack.extend(more.iter().cloned());
            }
        }
        false
    })
}

fn note_class_attrs(name: &str, attrs: std::collections::HashSet<String>) {
    if !name.is_empty() {
        PY_CLASS_ATTRS.with(|m| {
            m.borrow_mut().insert(name.to_string(), attrs);
        });
    }
}

fn class_has_attr(name: &str, attr: &str) -> bool {
    PY_CLASS_ATTRS.with(|attrs_map| {
        PY_CLASS_PARENTS.with(|parents_map| {
            let attrs_map = attrs_map.borrow();
            let parents_map = parents_map.borrow();
            let mut stack = vec![name.to_string()];
            let mut seen = std::collections::HashSet::new();
            while let Some(class_name) = stack.pop() {
                if !seen.insert(class_name.clone()) {
                    continue;
                }
                if attrs_map
                    .get(&class_name)
                    .map(|attrs| attrs.contains(attr))
                    .unwrap_or(false)
                {
                    return true;
                }
                if let Some(parents) = parents_map.get(&class_name) {
                    stack.extend(parents.iter().cloned());
                }
            }
            false
        })
    })
}

fn class_has_own_attr(name: &str, attr: &str) -> bool {
    PY_CLASS_ATTRS.with(|m| {
        m.borrow()
            .get(name)
            .map(|attrs| attrs.contains(attr))
            .unwrap_or(false)
    })
}

fn note_class_data_attrs(name: &str, attrs: std::collections::HashSet<String>) {
    if !name.is_empty() {
        PY_CLASS_DATA_ATTRS.with(|m| {
            m.borrow_mut().insert(name.to_string(), attrs);
        });
    }
}

fn class_has_data_attr(name: &str, attr: &str) -> bool {
    PY_CLASS_DATA_ATTRS.with(|m| {
        m.borrow()
            .get(name)
            .map(|attrs| attrs.contains(attr))
            .unwrap_or(false)
    })
}

fn note_instance_class(var: &str, class_name: &str) {
    PY_INSTANCE_CLASSES.with(|m| {
        m.borrow_mut().insert(var.to_string(), class_name.to_string());
    });
}

fn instance_class(var: &str) -> Option<String> {
    PY_INSTANCE_CLASSES.with(|m| m.borrow().get(var).cloned())
}

fn note_instance_attr(var: &str, attr: &str) {
    PY_INSTANCE_ATTRS.with(|m| {
        m.borrow_mut()
            .entry(var.to_string())
            .or_default()
            .insert(attr.to_string());
    });
}

fn instance_has_attr(var: &str, attr: &str) -> bool {
    PY_INSTANCE_ATTRS.with(|m| {
        m.borrow()
            .get(var)
            .map(|attrs| attrs.contains(attr))
            .unwrap_or(false)
    })
}

fn python_instance_index(var: &str, attr: &str) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(Expression::ident(var)),
        index: Box::new(Expression::string(attr)),
        null_safe: false })
}

fn is_userdict_instance(var: &str) -> bool {
    if is_userdict_var(var) {
        return true;
    }
    instance_class(var)
        .as_deref()
        .is_some_and(|class_name| py_class_is_subclass(class_name, "UserDict"))
}

fn userdict_data_expr(var: &str) -> Expression {
    python_instance_index(var, "data")
}

fn py_class_is_subclass(class_name: &str, target: &str) -> bool {
    if class_name == target || target == "object" {
        return true;
    }
    PY_CLASS_PARENTS.with(|m| {
        let parents = m.borrow();
        let mut stack = parents.get(class_name).cloned().unwrap_or_default();
        while let Some(parent) = stack.pop() {
            if parent == target || (parent == "bool" && target == "int") {
                return true;
            }
            if let Some(more) = parents.get(&parent) {
                stack.extend(more.iter().cloned());
            }
        }
        false
    })
}

thread_local! {
    // `Name = namedtuple('Type', 'f1 f2', defaults=[...])` records the type
    // name, ordered field names, and trailing defaults here. A later
    // `Name(args)` lowers to the shared `ExprKind::NamedTuple` (array-backed,
    // cross-language) via `call_or_new`. See `vybe_compiler::primitives::tuples`.
    static PY_NAMEDTUPLE_DEFS: std::cell::RefCell<std::collections::HashMap<String, NamedTupleDef>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
}

#[derive(Clone)]
struct NamedTupleDef {
    type_name: String,
    fields: Vec<String>,
    /// Trailing defaults (`namedtuple(..., defaults=[...])`); apply right-aligned.
    defaults: Vec<Expression> }

fn register_namedtuple_def(name: &str, def: NamedTupleDef) {
    PY_NAMEDTUPLE_DEFS.with(|m| {
        m.borrow_mut().insert(name.to_string(), def);
    });
}

fn namedtuple_def(name: &str) -> Option<NamedTupleDef> {
    PY_NAMEDTUPLE_DEFS.with(|m| m.borrow().get(name).cloned())
}

thread_local! {
    // Variables bound to a namedtuple instance (`p = P(1, 2)`), so `p._asdict()`
    // / `p._replace(...)` can desugar with the field names known statically.
    static PY_NAMEDTUPLE_INSTANCES: std::cell::RefCell<std::collections::HashMap<String, NamedTupleDef>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
}

fn record_namedtuple_instance(name: &str, def: NamedTupleDef) {
    PY_NAMEDTUPLE_INSTANCES.with(|m| {
        m.borrow_mut().insert(name.to_string(), def);
    });
}

fn namedtuple_instance_def(name: &str) -> Option<NamedTupleDef> {
    PY_NAMEDTUPLE_INSTANCES.with(|m| m.borrow().get(name).cloned())
}

/// The namedtuple definition backing a `_asdict`/`_replace` receiver: a direct
/// `NamedTuple` node (`P(1, 2)._replace(...)`) or a tracked instance variable
/// (`p = P(1, 2); p._replace(...)`).
fn receiver_namedtuple_def(recv: &Expression) -> Option<NamedTupleDef> {
    match &recv.kind {
        ExprKind::NamedTuple { fields, type_name } => Some(NamedTupleDef {
            type_name: type_name.clone().unwrap_or_default(),
            fields: fields.iter().filter_map(|(n, _)| n.clone()).collect(),
            defaults: Vec::new() }),
        ExprKind::Ident(name) => namedtuple_instance_def(name),
        _ => None }
}

/// Positional read `recv[index]` off a namedtuple receiver.
fn namedtuple_index_read(recv: &Expression, index: usize) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(recv.clone()),
        index: Box::new(Expression::int(index as i64)),
        null_safe: false })
}

/// `nt._asdict()` → an ordered dict `{field: nt[i]}`.
fn build_namedtuple_asdict(recv: &Expression, def: &NamedTupleDef) -> Expression {
    let props = def
        .fields
        .iter()
        .enumerate()
        .map(|(i, f)| ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str(f.clone()))),
            value: namedtuple_index_read(recv, i) })
        .collect();
    Expression::new(ExprKind::Object(props))
}

/// `nt._replace(**kw)` → a new namedtuple: each field keeps `nt[i]` unless a
/// keyword override supplies it. Reuses the shared `NamedTuple` lowering.
fn build_namedtuple_replace(
    recv: &Expression,
    def: &NamedTupleDef,
    args: Vec<Argument>,
) -> Expression {
    let mut overrides: HashMap<String, Expression> = HashMap::new();
    for arg in args {
        if let Some(name) = arg.name {
            overrides.insert(name, arg.value);
        }
    }
    let fields = def
        .fields
        .iter()
        .enumerate()
        .map(|(i, f)| {
            let value = overrides
                .remove(f)
                .unwrap_or_else(|| namedtuple_index_read(recv, i));
            (Some(f.clone()), value)
        })
        .collect();
    Expression::new(ExprKind::NamedTuple {
        fields,
        type_name: Some(def.type_name.clone()) })
}

/// Extract the string value of a string-literal expression.
fn str_literal(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        _ => None }
}

/// Ordered value expressions of a list/tuple literal.
fn sequence_values(expr: &Expression) -> Option<Vec<Expression>> {
    match &expr.kind {
        ExprKind::Tuple(items) | ExprKind::Set(items) => Some(items.clone()),
        ExprKind::Array(items) if items.iter().all(|e| e.key.is_none() && !e.spread) => {
            Some(items.iter().map(|e| e.value.clone()).collect())
        }
        _ => None }
}

/// Parse a namedtuple field spec: a whitespace/comma-separated string
/// (`'x y'`, `'x, y'`) or a list/tuple of name strings.
fn parse_field_spec(expr: &Expression) -> Option<Vec<String>> {
    if let Some(s) = str_literal(expr) {
        return Some(
            s.split([',', ' ', '\t', '\n'])
                .filter(|t| !t.is_empty())
                .map(|t| t.to_string())
                .collect(),
        );
    }
    let values = sequence_values(expr)?;
    values.iter().map(str_literal).collect()
}

/// If `value` is a `namedtuple('Type', fieldspec, defaults=...)` call, extract
/// its definition.
fn parse_namedtuple_call(value: &Expression) -> Option<NamedTupleDef> {
    let ExprKind::Call { callee, args, .. } = &value.kind else {
        return None;
    };
    let ExprKind::Ident(fname) = &callee.kind else {
        return None;
    };
    if fname != "namedtuple" && fname != "NamedTuple" {
        return None;
    }
    let positional: Vec<&Expression> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| &a.value)
        .collect();
    if positional.len() < 2 {
        return None;
    }
    let type_name = str_literal(positional[0])?;
    let fields = parse_field_spec(positional[1])?;
    let defaults = args
        .iter()
        .find(|a| a.name.as_deref() == Some("defaults"))
        .and_then(|a| sequence_values(&a.value))
        .unwrap_or_default();
    Some(NamedTupleDef {
        type_name,
        fields,
        defaults })
}

/// The type object bound by `Name = namedtuple(...)` — an object exposing
/// `_fields` (the field-name tuple) and `__typename`, so `Name._fields`
/// resolves as an ordinary member read.
fn namedtuple_type_object(def: &NamedTupleDef) -> Expression {
    let field_tuple = Expression::new(ExprKind::Tuple(
        def.fields
            .iter()
            .map(|f| Expression::new(ExprKind::Lit(Literal::Str(f.clone()))))
            .collect(),
    ));
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("_fields"),
            value: field_tuple },
        ObjectProperty::KeyValue {
            key: Expression::string("__typename"),
            value: Expression::new(ExprKind::Lit(Literal::Str(def.type_name.clone()))) },
    ]))
}

/// Build a `namedtuple` construction: bind positional/keyword args to fields in
/// order, filling trailing gaps from right-aligned `defaults`.
fn build_namedtuple_construction(def: &NamedTupleDef, args: Vec<Argument>) -> ExprKind {
    let n = def.fields.len();
    let mut values: Vec<Option<Expression>> = vec![None; n];
    let mut pos = 0usize;
    for arg in args {
        if let Some(name) = &arg.name {
            if let Some(i) = def.fields.iter().position(|f| f == name) {
                values[i] = Some(arg.value);
            }
        } else if pos < n {
            values[pos] = Some(arg.value);
            pos += 1;
        }
    }
    // Right-aligned defaults fill any still-missing trailing fields.
    let default_start = n.saturating_sub(def.defaults.len());
    for (i, slot) in values.iter_mut().enumerate() {
        if slot.is_none() && i >= default_start {
            *slot = def.defaults.get(i - default_start).cloned();
        }
    }
    let fields = def
        .fields
        .iter()
        .zip(values)
        .map(|(name, value)| (Some(name.clone()), value.unwrap_or_else(Expression::null)))
        .collect();
    ExprKind::NamedTuple {
        fields,
        type_name: Some(def.type_name.clone()) }
}

fn is_defined_class(name: &str) -> bool {
    PY_DEFINED_CLASSES.with(|m| m.borrow().contains(name))
}

/// Build a call expression, normalising `ClassName(args)` (a call whose callee
/// is a declared class) to `ExprKind::New` so construction has one canonical
/// shape across languages. Any other callee stays a plain `Call`.
/// The builtin a relational operator lowers to, mirroring how `+`/`-`
/// route through `__pyadd__`/`__pysub__` for the same reason: only a
/// Python-level helper can dispatch on an object operand.
fn py_relational_helper(op: BinOp) -> Option<&'static str> {
    Some(match op {
        BinOp::Lt => "__pylt__",
        BinOp::Gt => "__pygt__",
        BinOp::LtEq => "__pyle__",
        BinOp::GtEq => "__pyge__",
        _ => return None })
}

fn py_richcompare_method(op: BinOp) -> Option<&'static str> {
    Some(match op {
        BinOp::Lt => "__lt__",
        BinOp::Gt => "__gt__",
        BinOp::LtEq => "__le__",
        BinOp::GtEq => "__ge__",
        _ => return None })
}

fn py_fresh_class_lacks_richcompare(expr: &Expression, op: BinOp) -> bool {
    let Some(method) = py_richcompare_method(op) else {
        return false;
    };
    let ExprKind::New { class, .. } = &expr.kind else {
        return false;
    };
    let ExprKind::Ident(class_name) = &class.kind else {
        return false;
    };
    !class_has_attr(class_name, method)
}

/// `strftime` directives this expands, as `(property, pad width)`.
fn strftime_directive(spec: char) -> Option<(&'static str, i64)> {
    Some(match spec {
        'Y' => ("year", 4),
        'm' => ("month", 2),
        'd' => ("day", 2),
        'H' => ("hour", 2),
        'M' => ("minute", 2),
        'S' => ("second", 2),
        _ => return None })
}

/// `dt.strftime('%Y-%m-%d')` → `pad(dt['year'],4) + '-' + pad(dt['month'],2) …`
///
/// The components are already properties and the format is a literal, so
/// the whole format expands at compile time — no runtime format scanner,
/// and no host fn for it. Returns `None` for a non-literal format or a
/// directive outside the set above, so the call is left alone rather than
/// formatted wrongly.
fn strftime_expand(callee: &Expression, args: &[Argument]) -> Option<ExprKind> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "strftime" || args.len() != 1 {
        return None;
    }
    let ExprKind::Lit(Literal::Str(fmt)) = &args[0].value.kind else {
        return None;
    };

    let read = |prop: &str, width: i64| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__py_dt_pad")),
            args: vec![
                Argument::positional(Expression::new(ExprKind::Index {
                    object: Box::new((**object).clone()),
                    index: Box::new(Expression::string(prop)),
                    null_safe: false })),
                Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(width)))),
            ],
            optional: false })
    };

    let mut parts: Vec<Expression> = Vec::new();
    let mut lit = String::new();
    let mut chars = fmt.chars().peekable();
    while let Some(c) = chars.next() {
        if c != '%' {
            lit.push(c);
            continue;
        }
        match chars.next() {
            Some('%') => lit.push('%'),
            Some(spec) => {
                let (prop, width) = strftime_directive(spec)?;
                if !lit.is_empty() {
                    parts.push(Expression::string(&lit));
                    lit.clear();
                }
                parts.push(read(prop, width));
            }
            None => lit.push('%') }
    }
    if !lit.is_empty() {
        parts.push(Expression::string(&lit));
    }

    let mut iter = parts.into_iter();
    let first = iter.next().unwrap_or_else(|| Expression::string(""));
    Some(
        iter.fold(first, |acc, part| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(acc),
                right: Box::new(part) })
        })
        .kind,
    )
}

/// `time.strftime(fmt, t)` — same compile-time expansion as `dt.strftime`, but
/// over a `struct_time`'s `tm_*` fields, and it is a two-arg module function
/// rather than a method. `%A` reads the Monday=0 weekday name; `%j` is the
/// zero-padded day of year; `%%` is a literal `%`.
fn time_strftime_expand(callee: &Expression, args: &[Argument]) -> Option<ExprKind> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let ExprKind::Ident(module) = &object.kind else {
        return None;
    };
    if resolve_module_alias(module).unwrap_or_else(|| module.clone()) != "time"
        || field != "strftime"
        || args.len() != 2
    {
        return None;
    }
    let ExprKind::Lit(Literal::Str(fmt)) = &args[0].value.kind else {
        return None;
    };
    let t = args[1].value.clone();

    let field_read = |prop: &str| {
        Expression::new(ExprKind::Index {
            object: Box::new(t.clone()),
            index: Box::new(Expression::string(prop)),
            null_safe: false })
    };
    let padded = |prop: &str, width: i64| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__py_dt_pad")),
            args: vec![
                Argument::positional(field_read(prop)),
                Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(width)))),
            ],
            optional: false })
    };
    // `%A` → `['Monday', …][tm_wday]`.
    let weekday_name = || {
        let names = [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
        ];
        Expression::new(ExprKind::Index {
            object: Box::new(Expression::new(ExprKind::Array(
                names
                    .iter()
                    .map(|n| ArrayElement {
                        value: Expression::string(n),
                        spread: false,
                        key: None,
                        by_ref: false })
                    .collect(),
            ))),
            index: Box::new(field_read("tm_wday")),
            null_safe: false })
    };

    let mut parts: Vec<Expression> = Vec::new();
    let mut lit = String::new();
    let mut chars = fmt.chars();
    while let Some(c) = chars.next() {
        if c != '%' {
            lit.push(c);
            continue;
        }
        let part = match chars.next() {
            Some('%') => {
                lit.push('%');
                continue;
            }
            Some('Y') => padded("tm_year", 4),
            Some('m') => padded("tm_mon", 2),
            Some('d') => padded("tm_mday", 2),
            Some('H') => padded("tm_hour", 2),
            Some('M') => padded("tm_min", 2),
            Some('S') => padded("tm_sec", 2),
            Some('j') => padded("tm_yday", 3),
            Some('A') => weekday_name(),
            _ => return None };
        if !lit.is_empty() {
            parts.push(Expression::string(&lit));
            lit.clear();
        }
        parts.push(part);
    }
    if !lit.is_empty() {
        parts.push(Expression::string(&lit));
    }

    let mut iter = parts.into_iter();
    let first = iter.next().unwrap_or_else(|| Expression::string(""));
    Some(
        iter.fold(first, |acc, part| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(acc),
                right: Box::new(part) })
        })
        .kind,
    )
}

/// The components a datetime `replace` accepts, in the order the emitter
/// reads them. `microsecond`/`tzinfo`/`fold` are accepted as keywords but
/// have no effect on the ms-resolution, UTC-only value.
const DT_REPLACE_PARAMS: &[&str] = &[
    "year",
    "month",
    "day",
    "hour",
    "minute",
    "second",
    "microsecond",
    "tzinfo",
    "fold",
];

/// `d.replace(year=2021)` → `__py_dt_replace(d, year, …, second)`.
///
/// `value_methods` are keyed by method name alone, so `replace` already
/// belongs to `str.replace`. The two are told apart by shape rather than by
/// receiver type: a datetime `replace` names its components
/// (`replace(year=…)`), while `str.replace(old, new)` is positional. A call
/// with no keywords is therefore left alone for the string path.
fn datetime_replace_call(callee: &Expression, args: &[Argument]) -> Option<ExprKind> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "replace" || args.is_empty() {
        return None;
    }
    if !args
        .iter()
        .all(|a| matches!(a.name.as_deref(), Some(n) if DT_REPLACE_PARAMS.contains(&n)))
    {
        return None;
    }
    let mut call_args = vec![Argument::positional((**object).clone())];
    for prop in &DT_REPLACE_PARAMS[..6] {
        let value = args
            .iter()
            .find(|a| a.name.as_deref() == Some(*prop))
            .map(|a| a.value.clone())
            .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
        call_args.push(Argument::positional(value));
    }
    Some(ExprKind::Call {
        callee: Box::new(Expression::ident("__py_dt_replace")),
        args: call_args,
        optional: false })
}

/// `zoneinfo` — `ZoneInfo(key)` is a value carrying its key, and
/// `available_timezones()` is the set of zones this runtime can actually
/// resolve. `ecma:date` is UTC-only (`getTimezoneOffset` is always 0) and
/// no tzdata is mounted, so UTC is genuinely the whole set — this reports
/// what the runtime does, it does not stub out a larger list.
///
/// Both are plain data, so the walker builds them directly.
fn zoneinfo_call(callee: &Expression, args: &[Argument]) -> Option<ExprKind> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if !is_from_imported_module("zoneinfo") && !is_imported_module("zoneinfo") {
        return None;
    }
    match name.as_str() {
        "ZoneInfo" if args.len() == 1 => Some(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("__type"),
                value: Expression::string("ZoneInfo") },
            ObjectProperty::KeyValue {
                key: Expression::string("key"),
                value: args[0].value.clone() },
        ])),
        "available_timezones" if args.is_empty() => Some(ExprKind::Call {
            callee: Box::new(Expression::ident("set")),
            args: vec![Argument::positional(Expression::new(ExprKind::Array(
                vec![ArrayElement {
                    value: Expression::string("UTC"),
                    spread: false,
                    key: None,
                    by_ref: false }],
            )))],
            optional: false }),
        _ => None }
}

/// CPython signatures for the `datetime` constructors that are normally
/// called with keywords (`timedelta(days=2)`). Keyword handling belongs to
/// the frontend, so the emitter only ever sees positional arguments in this
/// exact order.
fn datetime_kwarg_signature(callee: &Expression) -> Option<&'static [&'static str]> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    let module = resolve_module_alias(name).unwrap_or_else(|| name.clone());
    if module != "datetime" {
        return None;
    }
    Some(match field.as_str() {
        "timedelta" => &[
            "days",
            "seconds",
            "microseconds",
            "milliseconds",
            "minutes",
            "hours",
            "weeks",
        ],
        "date" => &["year", "month", "day"],
        "time" => &["hour", "minute", "second", "microsecond", "tzinfo"],
        "datetime" => &[
            "year",
            "month",
            "day",
            "hour",
            "minute",
            "second",
            "microsecond",
            "tzinfo",
        ],
        _ => return None })
}

/// Place each keyword argument at its signature slot, filling the gaps with
/// `0`. Only runs when a keyword is actually present, so positional calls
/// keep their natural arity and the emitter's own defaults still apply.
fn normalize_datetime_kwargs(params: &[&str], args: Vec<Argument>) -> Vec<Argument> {
    if args.iter().all(|a| a.name.is_none()) {
        return args;
    }
    let zero = || Expression::new(ExprKind::Lit(Literal::Int(0)));
    let mut slots: Vec<Expression> = params.iter().map(|_| zero()).collect();
    let mut highest = 0usize;
    for (i, arg) in args.into_iter().enumerate() {
        let slot = match &arg.name {
            Some(name) => match params.iter().position(|p| p == name) {
                Some(pos) => pos,
                // An unknown keyword is not ours to interpret; drop it
                // rather than shift every later argument.
                None => continue },
            None => i };
        if slot < slots.len() {
            highest = highest.max(slot);
            slots[slot] = arg.value;
        }
    }
    slots
        .into_iter()
        .take(highest + 1)
        .map(Argument::positional)
        .collect()
}

/// Default `json.dumps` separators: `(", ", ": ")` compact, `(",", ": ")` when
/// `indent` is given (matching CPython's `json` module).
fn json_default_separators(has_indent: bool) -> (Expression, Expression) {
    if has_indent {
        (Expression::string(","), Expression::string(": "))
    } else {
        (Expression::string(", "), Expression::string(": "))
    }
}

/// Reshape `json.dumps(obj, cls=…, default=…, sort_keys=…, indent=…,
/// separators=…)` into the fixed positional form
/// `__py_json_dumps(value, default, sort_keys, indent, item_sep, kv_sep)` the
/// Maps an `os.path.FUNC` name to its pure-Python prelude helper, for the
/// string-math functions the walker re-routes off the profile/host path.
/// Is `e` the module `os` (directly or via an alias)?
fn is_os_module_ident(e: &Expression) -> bool {
    matches!(&e.kind, ExprKind::Ident(n)
        if n == "os" || resolve_module_alias(n).as_deref() == Some("os"))
}

/// `d.items()` → `[(p[0], p[1]) for p in __py_obj_entries__(d)]`. Python's
/// `dict.items()` yields (key, value) TUPLES, but `ecma:object.entries` returns
/// `[k, v]` LISTS. Building each pair as a tuple *literal* inside the
/// comprehension re-tags it (via the normal `ExprKind::Tuple` lowering), so
/// `list(d.items())` reprs as `[('a', 1)]` and `for k, v in d.items()` /
/// `dict(d.items())` still destructure the array backing. Entries is Map-aware,
/// so this is correct for the Map-backed dict.
fn rewrite_dict_items(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    if !args.is_empty() {
        return None;
    }
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "items" {
        return None;
    }
    let entries = call_ident("__py_obj_entries__", vec![(**object).clone()]);
    let pair_index = |i: i64| {
        Expression::new(ExprKind::Index {
            object: Box::new(Expression::new(ExprKind::Ident("__item_pair".into()))),
            index: Box::new(Expression::int(i)),
            null_safe: false })
    };
    let pair = Expression::new(ExprKind::Tuple(vec![pair_index(0), pair_index(1)]));
    Some(Expression::new(ExprKind::Comprehension {
        kind: ComprehensionKind::List,
        element: Box::new(pair),
        generators: vec![ComprehensionGen {
            target: Expression::new(ExprKind::Ident("__item_pair".into())),
            iter: entries,
            conditions: Vec::new(),
            is_async: false }] }))
}

/// `dict.fromkeys(keys[, value])` → `{__k: value for __k in keys}` (value
/// defaults to `None`). Reuses the dict-comprehension lowering (which builds a
/// Map), so the result is a real dict with the right keys/order — no separate
/// classmethod machinery needed.
/// `dict(...)` / `OrderedDict(...)` in a shape a dict LITERAL can express →
/// the literal node, which is what `{'a': 1}` already produces.
///
/// This is normalization, not lowering: the shared compiler carried an
/// `is_python_profile()` arm that rebuilt these three shapes by hand
/// (`primitives/calls.rs`), which also put the python-only NAMES `dict` and
/// `OrderedDict` in a shared crate. `OrderedDict` collapses to the same node
/// because ecma objects are insertion-ordered, which is exactly what that arm
/// said too.
///
/// Only the shapes a literal can represent are rewritten. `dict(other)`,
/// `dict(zip(a, b))` and a non-literal list argument fall through to the
/// ordinary call so the `dict` builtin still handles them.
fn rewrite_dict_construction(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if name != "dict" && name != "OrderedDict" && name != "Counter" {
        return None;
    }
    if args.iter().any(|arg| arg.spread || arg.by_ref) {
        return None;
    }

    // `Counter` is listed defensively and currently never arrives here — the
    // walker rewrites it upstream, so `Counter(a=1) + Counter(a=2)` keeps real
    // Counter semantics (verified byte-identical to python3). If that upstream
    // rewrite ever moves, this guard keeps the non-dict forms — iterable
    // `Counter([...])`, which COUNTS, and empty `Counter()` — off the literal
    // path, which would otherwise silently turn them into plain dicts.
    if name == "Counter" && (args.is_empty() || args.iter().any(|arg| arg.name.is_none())) {
        return None;
    }

    // `dict()` (vacuously all-named) and `dict(a=1, b=2)`.
    if args.iter().all(|arg| arg.name.is_some()) {
        let props = args
            .iter()
            .map(|arg| ObjectProperty::KeyValue {
                key: Expression::new(ExprKind::Lit(Literal::Str(
                    arg.name.clone().unwrap().into(),
                ))),
                value: arg.value.clone() })
            .collect();
        return Some(Expression::new(ExprKind::Object(props)));
    }

    // `dict([('x', 9), ('y', 8)])` — a LITERAL list of 2-tuples only.
    if args.len() == 1 && args[0].name.is_none() {
        if let ExprKind::Array(elements) = &args[0].value.kind {
            let mut props = Vec::with_capacity(elements.len());
            for element in elements {
                if element.spread || element.key.is_some() {
                    return None;
                }
                let ExprKind::Tuple(items) = &element.value.kind else {
                    return None;
                };
                if items.len() != 2 {
                    return None;
                }
                props.push(ObjectProperty::KeyValue {
                    key: items[0].clone(),
                    value: items[1].clone() });
            }
            return Some(Expression::new(ExprKind::Object(props)));
        }
    }

    None
}

fn rewrite_dict_fromkeys(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "fromkeys" || !matches!(&object.kind, ExprKind::Ident(n) if n == "dict") {
        return None;
    }
    if args.is_empty() || args.iter().any(|a| a.name.is_some() || a.spread) {
        return None;
    }
    let keys = args[0].value.clone();
    let value = args
        .get(1)
        .map(|a| a.value.clone())
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    // Dict-comprehension element is a 2-element Array [key, value] (the shape
    // the compiler unpacks), NOT a Tuple.
    let kv = |k: Expression, v: Expression| {
        Expression::new(ExprKind::Array(vec![
            ArrayElement { key: None, spread: false, by_ref: false, value: k },
            ArrayElement { key: None, spread: false, by_ref: false, value: v },
        ]))
    };
    Some(Expression::new(ExprKind::Comprehension {
        kind: ComprehensionKind::Dict,
        element: Box::new(kv(
            Expression::new(ExprKind::Ident("__fk_key".into())),
            value,
        )),
        generators: vec![ComprehensionGen {
            target: Expression::new(ExprKind::Ident("__fk_key".into())),
            iter: keys,
            conditions: Vec::new(),
            is_async: false }] }))
}

/// `random.NAME(args)` → `__py_random_NAME(args)` for the names that are not
/// reliably host-backed. `random`/`randint`/`choice`/`shuffle`/`sample`/`seed`
/// stay profile builtins. `randrange` and `choices` are arity-shaped here so
/// the prelude helpers stay fixed-arity (no default-binding assumptions).
fn rewrite_random_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let is_random = matches!(&object.kind, ExprKind::Ident(n)
        if n == "random" || resolve_module_alias(n).as_deref() == Some("random"));
    if !is_random {
        return None;
    }
    let call = |helper: &str, args: Vec<Argument>| {
        Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Ident(helper.to_string()))),
            args,
            optional: false }))
    };
    match field.as_str() {
        // Fixed-arity variates: forward positional args verbatim.
        "uniform" | "expovariate" | "gauss" | "normalvariate" | "lognormvariate"
        | "triangular" | "paretovariate" | "weibullvariate" | "vonmisesvariate"
        | "gammavariate" | "betavariate" | "getrandbits" | "randbytes" | "getstate"
        | "setstate" => call(&format!("__py_random_{field}"), args.to_vec()),
        // `randrange(stop)` / `(start, stop)` / `(start, stop, step)` → always
        // three positional args to the helper.
        "randrange" => {
            let pos: Vec<Expression> = args
                .iter()
                .filter(|a| a.name.is_none())
                .map(|a| a.value.clone())
                .collect();
            let (start, stop, step) = match pos.len() {
                1 => (Expression::int(0), pos[0].clone(), Expression::int(1)),
                2 => (pos[0].clone(), pos[1].clone(), Expression::int(1)),
                3 => (pos[0].clone(), pos[1].clone(), pos[2].clone()),
                _ => return None };
            call(
                "__py_random_randrange",
                vec![
                    Argument::positional(start),
                    Argument::positional(stop),
                    Argument::positional(step),
                ],
            )
        }
        // `choices(pop, weights=, cum_weights=, k=)` → four positional args.
        "choices" => {
            let kw = |name: &str| {
                args.iter()
                    .find(|a| a.name.as_deref() == Some(name))
                    .map(|a| a.value.clone())
            };
            let pop = args.iter().find(|a| a.name.is_none())?.value.clone();
            let weights = kw("weights").unwrap_or_else(Expression::null);
            let cum = kw("cum_weights").unwrap_or_else(Expression::null);
            let k = kw("k").unwrap_or_else(|| Expression::int(1));
            call(
                "__py_random_choices",
                vec![
                    Argument::positional(pop),
                    Argument::positional(weights),
                    Argument::positional(cum),
                    Argument::positional(k),
                ],
            )
        }
        _ => None }
}

/// json adapter consumes. Returns `None` for anything that isn't a
/// `json.dumps`-shaped call so normal handling proceeds.
fn rewrite_json_dumps(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "dumps" {
        return None;
    }
    if !matches!(&object.kind, ExprKind::Ident(n) if n == "json") {
        return None;
    }

    let value = args.iter().find(|a| a.name.is_none())?.value.clone();
    let kw = |name: &str| {
        args.iter()
            .find(|a| a.name.as_deref() == Some(name))
            .map(|a| a.value.clone())
    };

    // Encoder hook: `default=` wins, else `cls=` → `lambda __o: Cls().default(__o)`.
    let default_expr = if let Some(d) = kw("default") {
        d
    } else if let Some(cls) = kw("cls") {
        let inst = Expression::new(ExprKind::New {
            class: Box::new(cls),
            args: vec![] });
        let call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(inst),
                field: "default".into(),
                null_safe: false })),
            args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                "__o".into(),
            )))],
            optional: false });
        Expression::new(ExprKind::Lambda {
            params: vec![lambda_param("__o")],
            body: LambdaBody::Expr(Box::new(call)),
            is_async: false,
            captures: vec![] })
    } else {
        Expression::null()
    };

    let sort_keys =
        kw("sort_keys").unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(false))));
    let indent = kw("indent").unwrap_or_else(Expression::null);

    let (item_sep, kv_sep) = match kw("separators") {
        Some(sep) => match &sep.kind {
            ExprKind::Tuple(items) if items.len() == 2 => (items[0].clone(), items[1].clone()),
            _ => json_default_separators(kw("indent").is_some()) },
        None => json_default_separators(kw("indent").is_some()) };

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__py_json_dumps".into()))),
        args: vec![
            Argument::positional(value),
            Argument::positional(default_expr),
            Argument::positional(sort_keys),
            Argument::positional(indent),
            Argument::positional(item_sep),
            Argument::positional(kv_sep),
        ],
        optional: false }))
}

fn call_or_new(callee: Expression, args: Vec<Argument>) -> ExprKind {
    if let Some(kind) = zoneinfo_call(&callee, &args) {
        return kind;
    }
    if let Some(kind) = datetime_replace_call(&callee, &args) {
        return kind;
    }
    if let Some(kind) = strftime_expand(&callee, &args) {
        return kind;
    }
    if let Some(kind) = time_strftime_expand(&callee, &args) {
        return kind;
    }
    if let Some(params) = datetime_kwarg_signature(&callee) {
        let args = normalize_datetime_kwargs(params, args);
        return ExprKind::Call {
            callee: Box::new(callee),
            args,
            optional: false };
    }
    if let ExprKind::Ident(name) = &callee.kind {
        if let Some(def) = namedtuple_def(name) {
            return build_namedtuple_construction(&def, args);
        }
        if is_defined_class(name) {
            if is_callable_class(name) && !class_has_init(name) && !args.is_empty() {
                return ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::New {
                            class: Box::new(callee),
                            args: Vec::new() })),
                        field: "__call__".into(),
                        null_safe: false })),
                    args,
                    optional: false };
            }
            return ExprKind::New {
                class: Box::new(callee),
                args };
        }
    }
    // Inline `namedtuple('P', 'a b')(1, 2)` — the callee is the factory call.
    if let Some(def) = parse_namedtuple_call(&callee) {
        return build_namedtuple_construction(&def, args);
    }
    ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false }
}

/// Root identifier at the base of a member/index/call chain.
fn expr_root_ident(e: &Expression) -> Option<String> {
    match &e.kind {
        ExprKind::Ident(n) => Some(n.clone()),
        ExprKind::Member { object, .. } => expr_root_ident(object),
        ExprKind::Index { object, .. } => expr_root_ident(object),
        ExprKind::Call { callee, .. } => expr_root_ident(callee),
        _ => None }
}

/// True while an expression still denotes a module namespace: the import
/// root reached through nothing but attribute hops (`importlib.metadata`).
/// A `Call` or subscript anywhere in the chain produces an ordinary value,
/// so the namespace ends there.
fn is_module_namespace_path(e: &Expression) -> bool {
    module_namespace_path(e).is_some()
}

/// The dotted path of a module namespace expression, module-aliases
/// resolved (`md.version` where `import importlib.metadata as md` →
/// `importlib.metadata.version`). `None` once anything but an attribute
/// hop appears.
fn module_namespace_path(e: &Expression) -> Option<String> {
    match &e.kind {
        ExprKind::Ident(n) => {
            let module = resolve_module_alias(n).unwrap_or_else(|| n.clone());
            (is_imported_module(n) || py_module_surface(&module).is_some()).then_some(module)
        }
        ExprKind::Member { object, field, .. } => {
            Some(format!("{}.{}", module_namespace_path(object)?, field))
        }
        _ => None }
}

/// `calendar.month_name` / `calendar.day_name` — fixed, indexable name
/// tables. They are constants, so the walker materializes them directly
/// rather than routing a lookup through the emitter. `month_name[0]` is
/// empty because CPython's month numbering is 1-based.
fn calendar_name_table(path: &str) -> Option<&'static [&'static str]> {
    Some(match path {
        "calendar.month_name" => &[
            "",
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
        ],
        "calendar.month_abbr" => &[
            "", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ],
        "calendar.day_name" => &[
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
        ],
        "calendar.day_abbr" => &["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
        _ => return None })
}

/// The adapter's `__type` tag for a `datetime` type *reference* — the
/// second argument of `isinstance(x, datetime.date)`. Mirrors the tags
/// `emitter::datetime_adapter` stamps.
fn datetime_type_tag(e: &Expression) -> Option<&'static str> {
    let path = module_namespace_path(e)?;
    Some(match path.as_str() {
        "datetime.date" => "date",
        "datetime.time" => "time",
        "datetime.datetime" => "datetime",
        "datetime.timedelta" => "timedelta",
        "datetime.timezone" => "timezone",
        _ => return None })
}

/// `datetime` class attributes that hold *constructed* values rather than
/// scalars, which a `namespace_constants` entry cannot express. Each maps
/// to a zero-arg builtin the adapter materializes.
fn datetime_attr_builtin(path: &str) -> Option<&'static str> {
    Some(match path {
        "datetime.date.min" | "datetime.datetime.min" => "__py_date_min",
        "datetime.date.max" | "datetime.datetime.max" => "__py_date_max",
        "datetime.timezone.utc" => "__py_timezone_utc",
        "datetime.timedelta.resolution" => "__py_timedelta_resolution",
        _ => return None })
}

fn py_builtin_type_name(name: &str) -> Option<&'static str> {
    Some(match name {
        "int" => "int",
        "str" => "str",
        "list" => "list",
        "dict" => "dict",
        "tuple" => "tuple",
        "set" => "set",
        "bool" => "bool",
        "float" => "float",
        "bytes" => "bytes",
        "bytearray" => "bytearray",
        "complex" => "complex",
        "range" => "range",
        "object" => "object",
        "type" => "type",
        "NoneType" => "NoneType",
        "function" => "function",
        "builtin_function_or_method" => "builtin_function_or_method",
        "BaseException" => "BaseException",
        "Exception" => "Exception",
        "ArithmeticError" => "ArithmeticError",
        "AssertionError" => "AssertionError",
        "AttributeError" => "AttributeError",
        "EOFError" => "EOFError",
        "ExceptionGroup" => "ExceptionGroup",
        "FileNotFoundError" => "FileNotFoundError",
        "GeneratorExit" => "GeneratorExit",
        "GetoptError" => "GetoptError",
        "ImportError" => "ImportError",
        "IndexError" => "IndexError",
        "KeyboardInterrupt" => "KeyboardInterrupt",
        "KeyError" => "KeyError",
        "LookupError" => "LookupError",
        "NameError" => "NameError",
        "NotImplementedError" => "NotImplementedError",
        "OSError" => "OSError",
        "OverflowError" => "OverflowError",
        "RecursionError" => "RecursionError",
        "RuntimeError" => "RuntimeError",
        "StopIteration" => "StopIteration",
        "StopAsyncIteration" => "StopAsyncIteration",
        "SyntaxError" => "SyntaxError",
        "SystemExit" => "SystemExit",
        "TypeError" => "TypeError",
        "UnicodeError" => "UnicodeError",
        "ValueError" => "ValueError",
        "ZeroDivisionError" => "ZeroDivisionError",
        _ => return None })
}

fn py_builtin_exception_bases(name: &str) -> Option<&'static [&'static str]> {
    Some(match name {
        "BaseException" => &[],
        "Exception" => &["BaseException"],
        "ArithmeticError" => &["Exception", "BaseException"],
        "AssertionError" => &["Exception", "BaseException"],
        "AttributeError" => &["Exception", "BaseException"],
        "EOFError" => &["Exception", "BaseException"],
        "ExceptionGroup" => &["Exception", "BaseException"],
        "FileNotFoundError" => &["OSError", "Exception", "BaseException"],
        "GeneratorExit" => &["BaseException"],
        "GetoptError" => &["Exception", "BaseException"],
        "ImportError" => &["Exception", "BaseException"],
        "KeyboardInterrupt" => &["BaseException"],
        "LookupError" => &["Exception", "BaseException"],
        "NameError" => &["Exception", "BaseException"],
        "NotImplementedError" => &["RuntimeError", "Exception", "BaseException"],
        "OSError" => &["Exception", "BaseException"],
        "RuntimeError" => &["Exception", "BaseException"],
        "RecursionError" => &["RuntimeError", "Exception", "BaseException"],
        "StopIteration" => &["Exception", "BaseException"],
        "StopAsyncIteration" => &["Exception", "BaseException"],
        "SyntaxError" => &["Exception", "BaseException"],
        "SystemExit" => &["BaseException"],
        "TypeError" => &["Exception", "BaseException"],
        "UnicodeError" => &["ValueError", "Exception", "BaseException"],
        "ValueError" => &["Exception", "BaseException"],
        "IndexError" => &["LookupError", "Exception", "BaseException"],
        "KeyError" => &["LookupError", "Exception", "BaseException"],
        "OverflowError" => &["ArithmeticError", "Exception", "BaseException"],
        "ZeroDivisionError" => &["ArithmeticError", "Exception", "BaseException"],
        _ => return None })
}

fn py_builtin_exception_names() -> &'static [&'static str] {
    &[
        "BaseException",
        "Exception",
        "ArithmeticError",
        "AssertionError",
        "AttributeError",
        "EOFError",
        "ExceptionGroup",
        "FileNotFoundError",
        "GeneratorExit",
        "GetoptError",
        "ImportError",
        "IndexError",
        "KeyboardInterrupt",
        "KeyError",
        "LookupError",
        "NameError",
        "NotImplementedError",
        "OSError",
        "OverflowError",
        "RecursionError",
        "RuntimeError",
        "StopIteration",
        "StopAsyncIteration",
        "SyntaxError",
        "SystemExit",
        "TypeError",
        "UnicodeError",
        "ValueError",
        "ZeroDivisionError",
    ]
}

fn py_builtin_exception_ctor(name: &str) -> Option<String> {
    py_builtin_exception_bases(name)?;
    Some(format!("__py_exc_{name}"))
}

fn py_type_call_arg(e: &Expression) -> Option<&Expression> {
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(n) if n == "type") || args.len() != 1 {
        return None;
    }
    Some(&args[0].value)
}

fn py_static_type_name(e: &Expression) -> Option<&'static str> {
    match &e.kind {
        ExprKind::Lit(Literal::Bool(_)) => Some("bool"),
        ExprKind::Lit(Literal::Int(_)) => Some("int"),
        ExprKind::Lit(Literal::Float(_)) => Some("float"),
        ExprKind::Lit(Literal::Str(_)) => Some("str"),
        ExprKind::Lit(Literal::Bytes(_)) => Some("bytes"),
        ExprKind::Lit(Literal::Null) => Some("NoneType"),
        ExprKind::Array(_) => Some("list"),
        ExprKind::Tuple(_) | ExprKind::NamedTuple { .. } => Some("tuple"),
        ExprKind::Object(_)
        | ExprKind::Comprehension {
            kind: ComprehensionKind::Dict,
            ..
        } => Some("dict"),
        ExprKind::Set(_) => Some("set"),
        ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => Some("function"),
        ExprKind::New { class, .. } => {
            if let ExprKind::Ident(name) = &class.kind {
                py_builtin_type_name(name)
            } else {
                None
            }
        }
        ExprKind::Ident(name) if is_defined_class(name) => Some("type"),
        ExprKind::Ident(name) if is_defined_function(name) => Some("function"),
        ExprKind::Ident(name) if py_builtin_callable_lambda(name).is_some() => {
            Some("builtin_function_or_method")
        }
        ExprKind::Index { object, index, .. } => {
            if let ExprKind::New { class, .. } = &object.kind
                && let ExprKind::Ident(class_name) = &class.kind
                && let ExprKind::Lit(Literal::Str(field)) = &index.kind
                && class_has_attr(class_name, field)
                && !class_has_data_attr(class_name, field)
            {
                Some("method")
            } else {
                None
            }
        }
        ExprKind::Ident(name) => py_builtin_type_name(name).map(|_| "type"),
        ExprKind::Call { callee, args, .. }
            if matches!(&callee.kind, ExprKind::Ident(n) if n == "type") && args.len() == 1 =>
        {
            Some("type")
        }
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(n) if n.starts_with("__py_exc_") => {
                py_builtin_type_name(n.trim_start_matches("__py_exc_"))
            }
            ExprKind::Ident(n) if matches!(n.as_str(), "set" | "frozenset") => Some("set"),
            ExprKind::Ident(n) if n == "range" => Some("range"),
            ExprKind::Ident(n) if n == "bytes" || n == "__py_bytes_new__" => Some("bytes"),
            ExprKind::Ident(n) if n == "bytearray" => Some("bytearray"),
            ExprKind::Ident(n) if n == "complex" => Some("complex"),
            _ => None },
        _ => None }
}

fn py_type_is_builtin(left: &Expression, right: &Expression) -> Option<bool> {
    let value = py_type_call_arg(left)?;
    let ExprKind::Ident(type_name) = &right.kind else {
        return None;
    };
    Some(py_static_type_name(value)? == py_builtin_type_name(type_name)?)
}

fn py_builtin_subclass(sub: &str, base: &str) -> Option<bool> {
    py_builtin_type_name(sub)?;
    py_builtin_type_name(base)?;
    if py_builtin_exception_bases(sub).is_some() || py_builtin_exception_bases(base).is_some() {
        if sub == base {
            return Some(true);
        }
        let bases = py_builtin_exception_bases(sub)?;
        return Some(
            bases.contains(&base)
                || bases
                    .iter()
                    .any(|parent| py_builtin_subclass(parent, base) == Some(true)),
        );
    }
    Some(match (sub, base) {
        (a, b) if a == b => true,
        (_, "object") => true,
        ("bool", "int") => true,
        _ => false })
}

fn py_id_call_arg(e: &Expression) -> Option<&Expression> {
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(n) if n == "id") || args.len() != 1 {
        return None;
    }
    Some(&args[0].value)
}

fn py_fresh_object_expr(e: &Expression) -> bool {
    matches!(
        e.kind,
        ExprKind::Array(_)
            | ExprKind::Tuple(_)
            | ExprKind::Object(_)
            | ExprKind::Set(_)
            | ExprKind::New { .. }
    )
}

fn py_getattr_call_parts(e: &Expression) -> Option<(&Expression, &str)> {
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(n) if n == "getattr") || args.len() < 2 {
        return None;
    }
    let ExprKind::Lit(Literal::Str(attr)) = &args[1].value.kind else {
        return None;
    };
    Some((&args[0].value, attr))
}

fn py_static_getattr_member_identity(left: &Expression, right: &Expression) -> Option<bool> {
    let (obj, attr) = py_getattr_call_parts(left)?;
    let (object, field): (&Expression, &str) = match &right.kind {
        ExprKind::Member { object, field, .. } => (object, field.as_str()),
        ExprKind::Index { object, index, .. } => {
            let ExprKind::Lit(Literal::Str(field)) = &index.kind else {
                return None;
            };
            (object, field.as_str())
        }
        _ => return None };
    if field != attr {
        return Some(false);
    }
    let ExprKind::Ident(type_name) = &object.kind else {
        return None;
    };
    if let ExprKind::Ident(obj_name) = &obj.kind {
        if obj_name == type_name && py_builtin_type_name(type_name).is_some() {
            return Some(true);
        }
    }
    if py_static_type_name(obj) == py_builtin_type_name(type_name) {
        return Some(false);
    }
    None
}

fn py_static_callable(e: &Expression) -> Option<bool> {
    match &e.kind {
        ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => Some(true),
        ExprKind::Ident(name) if is_defined_class(name) => Some(true),
        ExprKind::Ident(name)
            if py_builtin_callable_lambda(name).is_some()
                || matches!(
                    name.as_str(),
                    "print"
                        | "len"
                        | "type"
                        | "isinstance"
                        | "issubclass"
                        | "callable"
                        | "int"
                        | "str"
                        | "list"
                        | "dict"
                        | "tuple"
                        | "set"
                        | "bool"
                        | "float"
                        | "bytes"
                        | "range"
                ) =>
        {
            Some(true)
        }
        ExprKind::New { class, .. } => {
            if let ExprKind::Ident(name) = &class.kind {
                Some(is_callable_class(name))
            } else {
                None
            }
        }
        _ => None }
}

fn py_static_hasattr(obj: &Expression, attr: &str) -> Option<bool> {
    if py_known_generator_expr(obj) {
        return Some(matches!(
            attr,
            "__iter__"
                | "__next__"
                | "send"
                | "throw"
                | "close"
                | "gi_frame"
                | "gi_running"
                | "gi_yieldfrom"
                | "__code__"
        ));
    }
    if let Some(type_name) = py_static_type_name(obj) {
        return Some(match (type_name, attr) {
            ("list", "append" | "extend" | "pop" | "sort" | "reverse" | "__len__") => true,
            ("tuple", "__len__") => true,
            ("dict", "keys" | "values" | "items" | "get" | "pop" | "__len__") => true,
            ("set", "add" | "discard" | "remove" | "__len__") => true,
            ("str", "upper" | "lower" | "replace" | "split" | "join" | "__len__") => true,
            ("int", "real") => true,
            _ => false });
    }
    if let ExprKind::New { class, .. } = &obj.kind {
        if let ExprKind::Ident(name) = &class.kind {
            return Some(class_has_attr(name, attr));
        }
    }
    None
}

/// `os.path.CONST` string/None constants (member reads, not calls). POSIX
/// values, matching CPython's `posixpath` module attributes.
fn os_path_constant(field: &str) -> Option<Expression> {
    Some(match field {
        "sep" => Expression::string("/"),
        "altsep" => Expression::null(),
        "pathsep" => Expression::string(":"),
        "extsep" => Expression::string("."),
        "curdir" => Expression::string("."),
        "pardir" => Expression::string(".."),
        "defpath" => Expression::string("/bin:/usr/bin"),
        "devnull" => Expression::string("/dev/null"),
        _ => return None })
}

/// Rewrite bare attribute reads to subscripts (see the module note above).
fn desugar_member_reads(e: Expression) -> Expression {
    match e.kind {
        ExprKind::Member {
            object,
            field,
            null_safe } => {
            // `string.<const>` — module constants (ascii_letters, digits, …).
            if matches!(&object.kind, ExprKind::Ident(n) if n == "string")
                && is_imported_module("string")
            {
                if let Some(lit) = string_module_constant(&field) {
                    return Expression::new(ExprKind::Lit(lit));
                }
                if let Some(name) = string_module_member(&field) {
                    return Expression::new(ExprKind::Ident(name.into()));
                }
            }
            // `stat.S_I*` / `stat.ST_*` integer constants.
            if matches!(&object.kind, ExprKind::Ident(n) if n == "stat")
                && is_imported_module("stat")
            {
                if let Some(lit) = stat_module_constant(&field) {
                    return Expression::new(ExprKind::Lit(lit));
                }
            }
            // `keyword.kwlist` / `keyword.softkwlist` are static interpreter data.
            if matches!(&object.kind, ExprKind::Ident(n) if n == "keyword")
                && is_imported_module("keyword")
            {
                if let Some(value) = keyword_module_member(&field) {
                    return value;
                }
            }
            // `mimetypes.types_map` / `encodings_map` / `suffix_map`.
            if module_namespace_path(&object).as_deref() == Some("mimetypes") {
                if let Some(value) = mimetypes_module_member(&field) {
                    return value;
                }
            }
            if module_namespace_path(&object).as_deref() == Some("getopt")
                && matches!(field.as_str(), "GetoptError" | "error")
            {
                return Expression::ident("GetoptError");
            }
            // `sys.<const>` scalars (platform, maxsize, byteorder, …).
            if matches!(&object.kind, ExprKind::Ident(n) if n == "sys") {
                if let Some(lit) = sys_module_constant(&field) {
                    return Expression::new(ExprKind::Lit(lit));
                }
                // `sys.version_info` is a TUPLE, not a scalar, so it cannot go
                // through `sys_module_constant`. Version-gated code compares it
                // (`sys.version_info >= (3, 8)`), which needs the lexicographic
                // sequence ordering in `emit_seq_relational`.
                if field == "version_info" {
                    return Expression::new(ExprKind::Tuple(vec![
                        Expression::int(3),
                        Expression::int(12),
                        Expression::int(0),
                        Expression::string("final"),
                        Expression::int(0),
                    ]));
                }
            }
            // `sqlite3.<const>` — DB-API module constants (static mount).
            if matches!(&object.kind, ExprKind::Ident(n) if n == "sqlite3") {
                if let Some(lit) = sqlite3_module_constant(&field) {
                    return Expression::new(ExprKind::Lit(lit));
                }
                // `sqlite3.Row` — a truthy callable row-factory sentinel. When a
                // connection's `row_factory` is set to it, fetch returns the raw
                // column-keyed row (named `row['k']` + positional `row[0]`).
                if field == "Row" {
                    return sqlite3_row_factory_lambda();
                }
            }
            // `os.path.CONST` — resolve to the POSIX literal before the object
            // chain is desugared (the `os.path` prefix is not a data read).
            if let ExprKind::Member {
                object: inner,
                field: pfield,
                ..
            } = &object.kind
            {
                if pfield == "path" && is_os_module_ident(inner) {
                    if let Some(lit) = os_path_constant(&field) {
                        return lit;
                    }
                }
            }
            let mut object = desugar_member_reads(*object);
            // Module-alias substitution: `m.dumps` where `m = json` compiles
            // exactly like `json.dumps` (profile builtins + ns resolution).
            if let ExprKind::Ident(n) = &object.kind {
                if let Some(module) = resolve_module_alias(n) {
                    object = Expression::new(ExprKind::Ident(module));
                }
            }
            if matches!(object.kind, ExprKind::Lit(Literal::Null)) && !null_safe {
                return py_raise_expr("AttributeError", Some("'NoneType' object has no attribute"));
            }
            if field == "__name__"
                && let ExprKind::Ident(name) = &object.kind
                && is_generator_func(name)
            {
                return Expression::string(name);
            }
            if py_known_generator_expr(&object) {
                match field.as_str() {
                    "__name__" => {
                        if let Some(name) = py_generator_expr_name(&object) {
                            return Expression::string(&name);
                        }
                    }
                    "__code__" => {
                        let name = py_generator_expr_name(&object).unwrap_or_default();
                        return Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                            key: Expression::string("co_name"),
                            value: Expression::string(&name) }]));
                    }
                    "gi_running" => return Expression::bool(false),
                    "gi_frame" | "gi_yieldfrom" => {
                        return Expression::new(ExprKind::Object(Vec::new()));
                    }
                    _ => {}
                }
            }
            if let ExprKind::Ident(var) = &object.kind
                && instance_has_attr(var, &field)
            {
                return python_instance_index(var, &field);
            }
            match field.as_str() {
                "real" | "numerator" => return object,
                "imag" => return Expression::int(0),
                "denominator" => return Expression::int(1),
                _ => {}
            }
            if let ExprKind::Ident(var) = &object.kind
                && is_chainmap_var(var)
                && field == "parents"
            {
                return call_ident("__py_chainmap_parents", vec![Expression::ident(var)]);
            }
            if let ExprKind::Ident(var) = &object.kind
                && is_userlist_var(var)
                && field == "data"
            {
                return Expression::ident(var);
            }
            if let ExprKind::Ident(var) = &object.kind
                && is_chainmap_var(var)
                && field == "maps"
            {
                return call_ident("__py_chainmap_maps", vec![Expression::ident(var)]);
            }
            if let ExprKind::Ident(var) = &object.kind
                && let Some(class_name) = instance_class(var)
                && class_has_data_attr(&class_name, &field)
                && !instance_has_attr(var, &field)
                && !in_assignment_target()
            {
                return Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&class_name)),
                    field,
                    null_safe });
            }
            // `types.ModuleType.__name__` — static metadata of the mounted
            // types surface.
            if field == "__name__" {
                if let Some(value) = py_type_call_arg(&object) {
                    if let Some(name) = py_static_type_name(value) {
                        return Expression::string(name);
                    }
                }
                if let ExprKind::Ident(name) = &object.kind {
                    if let Some(type_name) = py_builtin_type_name(name) {
                        return Expression::string(type_name);
                    }
                }
                if let ExprKind::Member {
                    object: inner_obj,
                    field: inner_field,
                    ..
                } = &object.kind
                {
                    if matches!(&inner_obj.kind, ExprKind::Ident(n) if n == "types")
                        && inner_field == "ModuleType"
                    {
                        return Expression::new(ExprKind::Lit(Literal::Str("type".into())));
                    }
                }
            }
            let root = expr_root_ident(&object);
            // `sys.modules` — the runtime mount registry. The walker knows
            // exactly which modules this unit mounts (PY_IMPORTED_MODULES),
            // so materialize the registry as a dict of alias → namespace
            // object (`'json' in sys.modules`, `sys.modules['json']`).
            if matches!(&object.kind, ExprKind::Ident(n) if n == "sys") && field == "modules" {
                if PY_SYS_MODULES_BOUND.with(|b| b.get()) {
                    return Expression::new(ExprKind::Ident("__py_sys_modules".into()));
                }
                let props: Vec<ObjectProperty> = PY_IMPORTED_MODULES.with(|m| {
                    m.borrow()
                        .iter()
                        .map(|name| ObjectProperty::KeyValue {
                            key: Expression::new(ExprKind::Lit(Literal::Str(name.clone().into()))),
                            value: Expression::new(ExprKind::Ident(name.clone())) })
                        .collect()
                });
                return Expression::new(ExprKind::Object(props));
            }
            // `importlib.reload` read as a value (`callable(importlib.reload)`)
            // — identity function over the module mount.
            if matches!(&object.kind, ExprKind::Ident(n) if n == "importlib") && field == "reload" {
                return Expression::new(ExprKind::Lambda {
                    params: vec![Param {
                        name: "__mod".into(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false }],
                    body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ident(
                        "__mod".into(),
                    )))),
                    is_async: false,
                    captures: vec![] });
            }
            // Module metadata resolves at COMPILE time — the walker knows
            // the mounts (§16.2 namespace bindings are compile-time):
            // `json.__name__` IS the import name; `__file__` is None for
            // host-backed component modules.
            if let ExprKind::Ident(module_name) = &object.kind {
                if is_imported_module(module_name) {
                    let module_name = resolve_module_alias(module_name)
                        .unwrap_or_else(|| module_name.clone());
                    if let Some(value) = dynamic_module_attr(&module_name, &field) {
                        return value;
                    }
                    if let Some(value) = py_module_metadata_attr(&module_name, &field) {
                        return value;
                    }
                    // `mod.__dict__` — a REAL Python dict built from the
                    // namespace object's entries, via the same dict-
                    // comprehension lowering `{p[0]: p[1] for p in
                    // Object.entries(mod)}` takes — so it carries the exact
                    // shape (`__keys`) every other dict has and
                    // `isinstance(x, dict)` / len / iteration behave
                    // identically.
                    if field == "__dict__" {
                        let entries = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(
                                "__py_obj_entries__".into(),
                            ))),
                            args: vec![Argument::positional(object)],
                            optional: false });
                        let pair_index = |i: i64| {
                            Expression::new(ExprKind::Index {
                                object: Box::new(Expression::new(ExprKind::Ident(
                                    "__py_dict_pair".into(),
                                ))),
                                index: Box::new(Expression::new(ExprKind::Lit(Literal::Int(i)))),
                                null_safe: false })
                        };
                        let element = Expression::new(ExprKind::Array(vec![
                            ArrayElement {
                                key: None,
                                spread: false,
                                by_ref: false,
                                value: pair_index(0) },
                            ArrayElement {
                                key: None,
                                spread: false,
                                by_ref: false,
                                value: pair_index(1) },
                        ]));
                        return Expression::new(ExprKind::Comprehension {
                            kind: ComprehensionKind::Dict,
                            element: Box::new(element),
                            generators: vec![ComprehensionGen {
                                target: Expression::new(ExprKind::Ident("__py_dict_pair".into())),
                                iter: entries,
                                conditions: Vec::new(),
                                is_async: false }] });
                    }
                }
            }
            if let Some(path) = module_namespace_path(&object) {
                if let Some(value) = dynamic_module_attr(&path, &field) {
                    return value;
                }
                if let Some(value) = py_module_metadata_attr(&path, &field) {
                    return value;
                }
            }
            // `datetime.date.min` / `datetime.timezone.utc` — class
            // attributes holding constructed values, which no scalar
            // constant entry can express.
            if let Some(path) = module_namespace_path(&object) {
                let full = format!("{path}.{field}");
                if let Some(builtin) = datetime_attr_builtin(&full) {
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident(builtin.into()))),
                        args: Vec::new(),
                        optional: false });
                }
                if let Some(names) = calendar_name_table(&full) {
                    return Expression::new(ExprKind::Array(
                        names
                            .iter()
                            .map(|n| ArrayElement {
                                value: Expression::new(ExprKind::Lit(Literal::Str((*n).into()))),
                                spread: false,
                                key: None,
                                by_ref: false })
                            .collect(),
                    ));
                }
            }
            // Keep `self.x` and `module.CONST` on the Member path. A module
            // read stays a namespace read only while the chain is still pure
            // attribute hops — once a call or subscript intervenes the result
            // is an ordinary value (`datetime.date(2020, 6, 15).year`) whose
            // attributes live on the data path, not the module surface.
            let keep = matches!(root.as_deref(), Some("self")) || is_module_namespace_path(&object);
            if keep {
                Expression::new(ExprKind::Member {
                    object: Box::new(object),
                    field,
                    null_safe })
            } else if in_assignment_target() {
                Expression::new(ExprKind::Index {
                    object: Box::new(object),
                    index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(field.into())))),
                    null_safe })
            } else {
                // An attribute READ is not a subscript. Both land in the same
                // map-backed storage, but they fail differently: `d["k"]`
                // raises KeyError, `o.k` raises AttributeError — and only an
                // attribute miss consults `__getattr__`. Lowering both to
                // `Index` threw that distinction away, so every attribute miss
                // surfaced the storage as a KeyError and the interceptor could
                // never run. Writes keep the Index form; the assignment path
                // has no miss to handle.
                call_ident(
                    "__py_attr_read",
                    vec![
                        object,
                        Expression::new(ExprKind::Lit(Literal::Str(field.into()))),
                    ],
                )
            }
        }
        ExprKind::Call {
            callee,
            args,
            optional } => {
            // `t.join()` — a THREAD join. Python's string/list `join` ALWAYS
            // takes the iterable as its argument (`sep.join(items)`), so a
            // zero-argument `.join()` is never the collection method and can
            // route to the shared thread join with no runtime check. Done here
            // rather than in the postfix walk because an empty argument list
            // does not reach the `call_args` rule at all.
            if args.is_empty()
                && let ExprKind::Member { object, field, .. } = &callee.kind
                && field == "join"
            {
                return call_ident(
                    "__py_thread_join",
                    vec![desugar_member_reads((**object).clone())],
                );
            }
            // `__import__('json')` — same static mount binding as
            // importlib.import_module.
            if let ExprKind::Ident(n) = &callee.kind {
                if let Some(rewritten) = collections_ctor_call(n, &args) {
                    return rewritten;
                }
                if n == "int"
                    && args.len() == 1
                    && let ExprKind::Lit(Literal::Str(s)) = &args[0].value.kind
                {
                    let trimmed = s.trim();
                    if trimmed.parse::<i64>().is_err() {
                        return py_raise_expr(
                            "ValueError",
                            Some(&format!("invalid literal for int() with base 10: '{}'", s)),
                        );
                    }
                }
                if n == "hash" && args.len() == 1 {
                    let value = desugar_member_reads(args[0].value.clone());
                    if matches!(value.kind, ExprKind::New { .. }) {
                        return Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(value),
                                field: "__hash__".into(),
                                null_safe: false })),
                            args: Vec::new(),
                            optional: false });
                    }
                }
                if n == "__import__" && args.len() == 1 {
                    if let Some(module_name) = resolve_string_const(&args[0].value) {
                        note_imported_module(&module_name);
                        return Expression::new(ExprKind::Ident(module_name));
                    }
                }
                // `getattr(module, 'lit')` — a static member read of the
                // mounted (and stamped) namespace object.
                if n == "getattr" && args.len() == 2 {
                    if let (ExprKind::Ident(m), ExprKind::Lit(Literal::Str(attr))) =
                        (&args[0].value.kind, &args[1].value.kind)
                    {
                        if is_imported_module(m) {
                            let module = resolve_module_alias(m).unwrap_or_else(|| m.clone());
                            if let Some(value) = dynamic_module_attr(&module, attr) {
                                return value;
                            }
                            if let Some(value) = py_module_metadata_attr(&module, attr) {
                                return value;
                            }
                            if let Some(value) = py_module_callable_member(&module, attr) {
                                return value;
                            }
                            if py_module_surface(&module)
                                .is_some_and(|surface| surface.iter().any(|name| *name == attr))
                                || py_module_renames(&module).is_some_and(|renames| {
                                    renames.iter().any(|(py, canon)| *py == attr || *canon == attr)
                                })
                            {
                                return Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::new(ExprKind::Ident(module))),
                                    field: attr.to_string(),
                                    null_safe: false });
                            }
                            return Expression::new(ExprKind::Index {
                                object: Box::new(Expression::new(ExprKind::Ident(module))),
                                index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                                    attr.to_string().into(),
                                )))),
                                null_safe: false });
                        }
                    }
                }
                // `hasattr(module, 'lit')` — the mounts are static, so the
                // answer is too (same category as find_spec/sys.modules).
                if n == "hasattr" && args.len() == 2 {
                    let module_path = match &args[0].value.kind {
                        // The alias pass may already have substituted the
                        // bound name with its dotted module path
                        // (`md` → `importlib.metadata`), so accept a known
                        // dotted surface directly too.
                        ExprKind::Ident(m)
                            if is_imported_module(m)
                                || resolve_module_alias(m).is_some()
                                || py_module_surface(m).is_some() =>
                        {
                            Some(resolve_module_alias(m).unwrap_or_else(|| m.clone()))
                        }
                        ExprKind::Member {
                            object: o,
                            field: f,
                            ..
                        } => match &o.kind {
                            ExprKind::Ident(m) if is_imported_module(m) => Some(format!("{m}.{f}")),
                            _ => None },
                        _ => None };
                    if let (Some(path), ExprKind::Lit(Literal::Str(attr))) =
                        (module_path, &args[1].value.kind)
                    {
                        // Module metadata dunders always exist on a module.
                        if matches!(
                            attr.as_ref(),
                            "__name__"
                                | "__file__"
                                | "__package__"
                                | "__doc__"
                                | "__loader__"
                                | "__spec__"
                        ) {
                            return Expression::new(ExprKind::Lit(Literal::Bool(true)));
                        }
                        if dynamic_module_attr(&path, attr.as_ref()).is_some() {
                            return Expression::new(ExprKind::Lit(Literal::Bool(true)));
                        }
                        if let Some(surface) = py_module_surface(&path) {
                            let attr_str: &str = attr.as_ref();
                            let has = surface.iter().any(|a| *a == attr_str);
                            return Expression::new(ExprKind::Lit(Literal::Bool(has)));
                        }
                    }
                }
            }
            // `pkgutil.iter_modules()` — no filesystem package walk in a
            // component world; the static answer is the empty list.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if matches!(&object.kind, ExprKind::Ident(n) if n == "inspect") {
                    match field.as_str() {
                        "isgenerator" if args.len() == 1 => {
                            return Expression::bool(py_known_generator_expr(&args[0].value));
                        }
                        "isgeneratorfunction" if args.len() == 1 => {
                            let is_gen_fn = matches!(
                                &args[0].value.kind,
                                ExprKind::Ident(name) if is_generator_func(name)
                            );
                            return Expression::bool(is_gen_fn);
                        }
                        "getgeneratorstate" if args.len() == 1 => {
                            if py_known_generator_expr(&args[0].value) {
                                return Expression::string("GEN_SUSPENDED");
                            }
                        }
                        _ => {}
                    }
                }
                if matches!(&object.kind, ExprKind::Ident(n) if n == "pkgutil")
                    && field == "iter_modules"
                {
                    let modules = ["json", "math", "os", "sys", "collections"];
                    return Expression::new(ExprKind::Array(
                        modules
                            .into_iter()
                            .map(|name| ArrayElement {
                                key: None,
                                spread: false,
                                by_ref: false,
                                value: Expression::new(ExprKind::Object(vec![
                                    ObjectProperty::KeyValue {
                                        key: Expression::new(ExprKind::Lit(Literal::Str(
                                            "name".into(),
                                        ))),
                                        value: Expression::new(ExprKind::Lit(Literal::Str(
                                            name.into(),
                                        ))) },
                                ])) })
                            .collect(),
                    ));
                }
                // `types.ModuleType('name')` — a module object with its
                // `__name__` metadata.
                if matches!(&object.kind, ExprKind::Ident(n) if n == "types")
                    && field == "ModuleType"
                    && args.len() == 1
                {
                    let name = args.into_iter().next().unwrap().value;
                    return Expression::new(ExprKind::Object(vec![
                        ObjectProperty::KeyValue {
                            key: Expression::new(ExprKind::Lit(Literal::Str("__name__".into()))),
                            value: name.clone() },
                        ObjectProperty::KeyValue {
                            key: Expression::new(ExprKind::Lit(Literal::Str("__file__".into()))),
                            value: Expression::string("<module>") },
                        ObjectProperty::KeyValue {
                            key: Expression::new(ExprKind::Lit(Literal::Str("__doc__".into()))),
                            value: Expression::string("") },
                        ObjectProperty::KeyValue {
                            key: Expression::new(ExprKind::Lit(Literal::Str("__spec__".into()))),
                            value: Expression::new(ExprKind::Object(vec![
                                ObjectProperty::KeyValue {
                                    key: Expression::new(ExprKind::Lit(Literal::Str(
                                        "name".into(),
                                    ))),
                                    value: name },
                                ObjectProperty::KeyValue {
                                    key: Expression::new(ExprKind::Lit(Literal::Str(
                                        "loader".into(),
                                    ))),
                                    value: Expression::new(ExprKind::Object(vec![])) },
                            ])) },
                    ]));
                }
            }
            // `importlib.import_module('json')` with a literal module name →
            // the mounted namespace-object global (same binding `import json`
            // reads). Registers the name so later `json.X` reads stay
            // namespace access.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if matches!(&object.kind, ExprKind::Ident(n) if n == "importlib")
                    && field == "import_module"
                    && args.len() == 1
                {
                    if let Some(module_name) = resolve_string_const(&args[0].value) {
                        note_imported_module(&module_name);
                        return Expression::new(ExprKind::Ident(module_name));
                    }
                }
                // `importlib.reload(m)` — modules are immutable mounts here;
                // reload is the identity per its contract (returns the module).
                if matches!(&object.kind, ExprKind::Ident(n) if n == "importlib")
                    && field == "reload"
                    && args.len() == 1
                {
                    return desugar_member_reads(args.into_iter().next().unwrap().value);
                }
                // `importlib.util.find_spec('json')` with a literal name →
                // compile-time spec object `{name: 'json'}` (mounts are
                // static; a found spec is truthy with a `.name`).
                if field == "find_spec" && args.len() == 1 {
                    if let ExprKind::Member {
                        object: inner_obj,
                        field: inner_field,
                        ..
                    } = &object.kind
                    {
                        if matches!(&inner_obj.kind, ExprKind::Ident(n) if n == "importlib")
                            && inner_field == "util"
                        {
                            if let Some(module_name) = resolve_string_const(&args[0].value) {
                                return py_module_spec_object(&module_name);
                            }
                        }
                    }
                }
            }
            // `list.sort(key=f[, reverse=True])` — the array sort ignores the
            // Python key, so route to the in-place key-sort primitive.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field == "sort" {
                    if let Some(key) = args
                        .iter()
                        .find(|a| a.name.as_deref() == Some("key"))
                        .map(|a| wrap_key_ident_in_lambda(a.value.clone()))
                    {
                        let recv = desugar_member_reads((**object).clone());
                        let sorted = call_ident("__py_sort_by_key", vec![recv, key]);
                        let out = if args.iter().any(|a| a.name.as_deref() == Some("reverse")) {
                            Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new(sorted),
                                    field: "reverse".into(),
                                    null_safe: false })),
                                args: vec![],
                                optional: false })
                        } else {
                            sorted
                        };
                        return out;
                    }
                }
            }
            // sqlite3: `sqlite3.connect(...)` and methods on tracked
            // connection/cursor variables → collision-free `__sql_*` builtins.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let ExprKind::Ident(name) = &object.kind {
                    if let Some(source) = mapping_proxy_source(name) {
                        let recv = desugar_member_reads(source);
                        return Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(recv),
                                field: field.clone(),
                                null_safe: false })),
                            args,
                            optional });
                    }
                    if let Some(settings) = textwrapper_args(name) {
                        let values: Vec<Expression> = args
                            .iter()
                            .map(|a| desugar_member_reads(a.value.clone()))
                            .collect();
                        if matches!(field.as_str(), "wrap" | "fill") && values.len() == 1 {
                            if let Some(folded) = fold_textwrapper_method(field, &settings, &args) {
                                return folded;
                            }
                            let mut call_args = Vec::with_capacity(settings.len() + 1);
                            call_args.push(values[0].clone());
                            call_args.extend(settings);
                            let helper = if field == "wrap" {
                                "__py_textwrap_wrap"
                            } else {
                                "__py_textwrap_fill"
                            };
                            return call_ident(helper, call_args);
                        }
                    }
                }
                if let Some(path) = module_namespace_path(object)
                    && path == "collections"
                    && let Some(rewritten) = collections_ctor_call(field, &args)
                {
                    return rewritten;
                }
                if let Some(path) = module_namespace_path(object)
                    && path == "fnmatch"
                    && let Some(folded) = fold_fnmatch_call(field, &args)
                {
                    return folded;
                }
                if matches!(&object.kind, ExprKind::Ident(n) if n == "Counter")
                    && field == "fromkeys"
                    && !args.is_empty()
                {
                    let keys = desugar_member_reads(args[0].value.clone());
                    let value = args
                        .iter()
                        .find(|a| a.name.as_deref() == Some("v"))
                        .map(|a| desugar_member_reads(a.value.clone()))
                        .or_else(|| args.get(1).map(|a| desugar_member_reads(a.value.clone())))
                        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                    return call_ident("__py_counter_fromkeys", vec![keys, value]);
                }
                if matches!(&object.kind, ExprKind::Ident(n) if n == "types")
                    && matches!(
                        field.as_str(),
                        "SimpleNamespace"
                            | "MappingProxyType"
                            | "MethodType"
                            | "DynamicClassAttribute"
                            | "new_class"
                            | "resolve_bases"
                    )
                {
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(field)),
                        args,
                        optional });
                }
                if let Some(path) = module_namespace_path(object) {
                    if let Some(value) = dynamic_module_attr(&path, field) {
                        let args = args
                            .into_iter()
                            .map(|mut a| {
                                a.value = desugar_member_reads(a.value);
                                a
                            })
                            .collect();
                        return Expression::new(ExprKind::Call {
                            callee: Box::new(value),
                            args,
                            optional });
                    }
                }
                if let Some(rewritten) =
                    rewrite_sqlite_call(object, field, args.clone(), optional)
                {
                    return rewritten;
                }
                if let Some(rewritten) = rewrite_mimetypes_call(object, field, &args) {
                    return rewritten;
                }
                if let Some(rewritten) = rewrite_getopt_call(object, field, &args) {
                    return rewritten;
                }
                // `stat.S_IMODE(m)` / `stat.S_ISDIR(m)` → inline bitwise exprs.
                if is_imported_module("stat") {
                    if let Some(rewritten) = rewrite_stat_call(object, field, &args) {
                        return rewritten;
                    }
                }
                // `sys.intern(s)` / `sys.getrecursionlimit()` etc.
                if let Some(rewritten) = rewrite_sys_call(object, field, &args) {
                    return rewritten;
                }
                // `keyword.iskeyword(s)` / `keyword.issoftkeyword(s)`.
                if is_imported_module("keyword") {
                    if let Some(rewritten) = rewrite_keyword_call(object, field, &args) {
                        return rewritten;
                    }
                }
                // `platform.system()` / `platform.uname()` etc.
                if is_imported_module("platform") {
                    if let Some(rewritten) = rewrite_platform_call(object, field) {
                        return rewritten;
                    }
                }
                // `html.escape(s)` / `html.unescape(s)`.
                if is_imported_module("html") {
                    if let Some(rewritten) = rewrite_html_call(object, field, &args) {
                        return rewritten;
                    }
                }
                // `re.search(...)` / `m.group(i)` over ecma:regexp.
                if is_imported_module("re") {
                    if let Some(rewritten) = rewrite_re_call(object, field, &args) {
                        return rewritten;
                    }
                    if let Some(rewritten) = rewrite_re_match_method(object, field, &args) {
                        return rewritten;
                    }
                }
                if let ExprKind::Index {
                    object: indexed_object,
                    index,
                    ..
                } = &object.kind
                    && let ExprKind::Ident(var) = &indexed_object.kind
                    && let Some(factory) = defaultdict_factory(var)
                {
                    let vals: Vec<Expression> = args
                        .iter()
                        .map(|a| desugar_member_reads(a.value.clone()))
                        .collect();
                    if field == "append" && vals.len() == 1 {
                        return call_ident(
                            "__py_defaultdict_append",
                            vec![Expression::ident(var), factory, *index.clone(), vals[0].clone()],
                        );
                    }
                    if field == "add" && vals.len() == 1 {
                        return call_ident(
                            "__py_defaultdict_add",
                            vec![Expression::ident(var), factory, *index.clone(), vals[0].clone()],
                        );
                    }
                }
                if let ExprKind::Call {
                    callee: indexed_callee,
                    args: indexed_args,
                    ..
                } = &object.kind
                    && matches!(&indexed_callee.kind, ExprKind::Ident(n) if n == "__py_defaultdict_get")
                    && indexed_args.len() == 3
                {
                    let vals: Vec<Expression> = args
                        .iter()
                        .map(|a| desugar_member_reads(a.value.clone()))
                        .collect();
                    let recv = indexed_args[0].value.clone();
                    let factory = indexed_args[1].value.clone();
                    let key = indexed_args[2].value.clone();
                    if field == "append" && vals.len() == 1 {
                        return call_ident(
                            "__py_defaultdict_append",
                            vec![recv, factory, key, vals[0].clone()],
                        );
                    }
                    if field == "add" && vals.len() == 1 {
                        return call_ident(
                            "__py_defaultdict_add",
                            vec![recv, factory, key, vals[0].clone()],
                        );
                    }
                }
                if let ExprKind::Ident(var) = &object.kind {
                    if let Some(maxlen) = deque_maxlen(var) {
                        let recv = desugar_member_reads((**object).clone());
                        let vals: Vec<Expression> = args
                            .iter()
                            .map(|a| desugar_member_reads(a.value.clone()))
                            .collect();
                        match field.as_str() {
                            "append" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_deque_append",
                                    vec![recv, vals[0].clone(), maxlen],
                                );
                            }
                            "appendleft" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_deque_appendleft",
                                    vec![recv, vals[0].clone(), maxlen],
                                );
                            }
                            "extend" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_deque_extend",
                                    vec![recv, vals[0].clone(), maxlen],
                                );
                            }
                            "extendleft" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_deque_extendleft",
                                    vec![recv, vals[0].clone(), maxlen],
                                );
                            }
                            "remove" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_deque_remove",
                                    vec![recv, vals[0].clone()],
                                );
                            }
                            _ => {}
                        }
                    }
                    if is_userlist_var(var) {
                        let recv = desugar_member_reads((**object).clone());
                        let vals: Vec<Expression> = args
                            .iter()
                            .map(|a| desugar_member_reads(a.value.clone()))
                            .collect();
                        match field.as_str() {
                            "append" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_deque_append",
                                    vec![
                                        recv,
                                        vals[0].clone(),
                                        Expression::new(ExprKind::Lit(Literal::Null)),
                                    ],
                                );
                            }
                            "extend" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_deque_extend",
                                    vec![
                                        recv,
                                        vals[0].clone(),
                                        Expression::new(ExprKind::Lit(Literal::Null)),
                                    ],
                                );
                            }
                            _ => {}
                        }
                    }
                    if is_counter_expr(object) {
                        let recv = desugar_member_reads((**object).clone());
                        let vals: Vec<Expression> = args
                            .iter()
                            .map(|a| desugar_member_reads(a.value.clone()))
                            .collect();
                        match field.as_str() {
                            "update" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_counter_update",
                                    vec![recv, vals[0].clone()],
                                );
                            }
                            "subtract" if vals.len() == 1 => {
                                return call_ident(
                                    "__py_counter_subtract",
                                    vec![recv, vals[0].clone()],
                                );
                            }
                            "elements" if vals.is_empty() => {
                                return call_ident("__py_counter_elements", vec![recv]);
                            }
                            "total" if vals.is_empty() => {
                                return call_ident("__py_counter_total", vec![recv]);
                            }
                            _ => {}
                        }
                    }
                    if is_chainmap_var(var) {
                        let recv = desugar_member_reads((**object).clone());
                        let vals: Vec<Expression> = args
                            .iter()
                            .map(|a| desugar_member_reads(a.value.clone()))
                            .collect();
                        if field == "new_child" {
                            let child = vals
                                .first()
                                .cloned()
                                .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                            return call_ident("__py_chainmap_new_child", vec![recv, child]);
                        }
                    }
                }
                if field == "most_common" && args.len() <= 1 {
                    let mut vals = vec![desugar_member_reads((**object).clone())];
                    vals.extend(args.iter().map(|a| desugar_member_reads(a.value.clone())));
                    return call_ident("__py_counter_most_common", vals);
                }
                if field == "move_to_end" && !args.is_empty() && args.len() <= 2 {
                    let recv = desugar_member_reads((**object).clone());
                    let key = desugar_member_reads(args[0].value.clone());
                    let last = args
                        .iter()
                        .find(|a| a.name.as_deref() == Some("last"))
                        .map(|a| desugar_member_reads(a.value.clone()))
                        .or_else(|| args.get(1).map(|a| desugar_member_reads(a.value.clone())))
                        .unwrap_or_else(|| Expression::bool(true));
                    return call_ident("__py_ordereddict_move_to_end", vec![recv, key, last]);
                }
            }
            // `operator.<fn>(...)` — see [operator_call_lowering] for why these
            // specific names cannot go through the profile. Arguments are
            // desugared first so member reads inside them (`operator.truth(o.x)`)
            // resolve the same as anywhere else.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && matches!(&object.kind, ExprKind::Ident(n) if n == "operator")
                && is_imported_module("operator")
            {
                let desugared: Vec<Argument> = args
                    .iter()
                    .cloned()
                    .map(|mut a| {
                        a.value = desugar_member_reads(a.value);
                        a
                    })
                    .collect();
                if let Some(lowered) = operator_call_lowering(field, &desugared) {
                    return lowered;
                }
            }
            // `tempfile.NamedTemporaryFile(prefix=…, suffix=…, dir=…)` etc. —
            // adapters see only a stack of values, never argument NAMES, so the
            // keywords are flattened here into a fixed (prefix, suffix, dir)
            // order with "" defaults.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && matches!(&object.kind, ExprKind::Ident(n) if n == "tempfile")
                && matches!(
                    field.as_str(),
                    "NamedTemporaryFile" | "TemporaryFile" | "TemporaryDirectory" | "mkdtemp"
                        | "mkstemp"
                )
            {
                let kw = |name: &str| {
                    args.iter()
                        .find(|a| a.name.as_deref() == Some(name))
                        .map(|a| desugar_member_reads(a.value.clone()))
                        .unwrap_or_else(|| Expression::string(""))
                };
                let fixed = vec![
                    Argument::positional(kw("prefix")),
                    Argument::positional(kw("suffix")),
                    Argument::positional(kw("dir")),
                ];
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("tempfile")),
                        field: field.clone(),
                        null_safe: false })),
                    args: fixed,
                    optional });
            }
            // `tempfile.NamedTemporaryFile(prefix=…, suffix=…, dir=…)` etc.
            // `emit_common(name, chunks, current, argc, line)` receives a value
            // stack and a COUNT — argument names do not survive to emit time —
            // so the keywords are flattened here into a fixed
            // (prefix, suffix, dir) order with "" defaults, the same way
            // `json.dumps(indent=…)` and `sorted(key=…)` are handled.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && matches!(&object.kind, ExprKind::Ident(n) if n == "tempfile")
                && matches!(
                    field.as_str(),
                    "NamedTemporaryFile"
                        | "TemporaryFile"
                        | "TemporaryDirectory"
                        | "mkdtemp"
                        | "mkstemp"
                )
            {
                let kw = |name: &str| {
                    args.iter()
                        .find(|a| a.name.as_deref() == Some(name))
                        .map(|a| desugar_member_reads(a.value.clone()))
                        .unwrap_or_else(|| Expression::string(""))
                };
                let fixed = vec![
                    Argument::positional(kw("prefix")),
                    Argument::positional(kw("suffix")),
                    Argument::positional(kw("dir")),
                ];
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("tempfile")),
                        field: field.clone(),
                        null_safe: false })),
                    args: fixed,
                    optional });
            }
            // `pprint.pformat(...)` / `pprint.PrettyPrinter(...)` / … — call the
            // injected prelude global (see [PPRINT_PRELUDE]). Kept a real Call so
            // keyword args (`pformat(d, width=20, sort_dicts=False)`) survive.
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && matches!(&object.kind, ExprKind::Ident(n) if n == "pprint")
                && let Some(name) = pprint_module_member(field)
            {
                let args = args
                    .into_iter()
                    .map(|mut a| {
                        a.value = desugar_member_reads(a.value);
                        a
                    })
                    .collect();
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
                    args,
                    optional });
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && let Some(path) = module_namespace_path(object)
            {
                let rewritten = match path.as_str() {
                    "shlex" => shlex_module_member(field),
                    "textwrap" => textwrap_module_member(field),
                    _ => None };
                if let Some(name) = rewritten {
                    let mut args: Vec<Argument> = args
                        .into_iter()
                        .map(|mut a| {
                            a.value = desugar_member_reads(a.value);
                            a
                        })
                        .collect();
                    if path == "shlex" {
                        args = normalize_shlex_call_args(field, args);
                    } else if path == "textwrap" {
                        args = flatten_textwrap_args(field, args);
                        if let Some(folded) = fold_textwrap_call(field, &args) {
                            return folded;
                        }
                    }
                    return if matches!(name, "__py_shlex_class" | "__py_TextWrapper") {
                        Expression::new(ExprKind::New {
                            class: Box::new(Expression::new(ExprKind::Ident(name.into()))),
                            args })
                    } else {
                        Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
                            args,
                            optional })
                    };
                }
            }
            // `string.Template(...)` / `string.Formatter()` / `string.capwords(...)`
            // — call the injected prelude global (see [STRING_PRELUDE]). Kept as a
            // real Call so keyword args (e.g. `capwords(s, sep="-")`) survive.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if matches!(&object.kind, ExprKind::Ident(n) if n == "functools")
                    && field == "wraps"
                {
                    let args = args
                        .into_iter()
                        .map(|mut a| {
                            a.value = desugar_member_reads(a.value);
                            a
                        })
                        .collect();
                    return Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("wraps")),
                        args,
                        optional });
                }
                if matches!(&object.kind, ExprKind::Ident(n) if n == "string")
                    && is_imported_module("string")
                {
                    if let Some(name) = string_module_member(field) {
                        let args = args
                            .into_iter()
                            .map(|mut a| {
                                a.value = desugar_member_reads(a.value);
                                a
                            })
                            .collect();
                        return Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
                            args,
                            optional });
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && let ExprKind::Ident(var) = &object.kind
                && is_userdict_instance(var)
                && matches!(field.as_str(), "keys" | "items" | "values" | "get")
            {
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(userdict_data_expr(var)),
                        field: field.clone(),
                        null_safe: false })),
                    args: args
                        .into_iter()
                        .map(|mut a| {
                            a.value = desugar_member_reads(a.value);
                            a
                        })
                        .collect(),
                    optional });
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && field == "split"
                && args.len() == 1
                && matches!(&args[0].value.kind, ExprKind::Ident(name) if py_builtin_exception_bases(name).is_some() || is_defined_class(name))
            {
                let ExprKind::Ident(type_name) = &args[0].value.kind else {
                    unreachable!();
                };
                return call_ident(
                    "__py_exception_group_split",
                    vec![
                        desugar_member_reads((**object).clone()),
                        Expression::string(type_name),
                    ],
                );
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind
                && field == "__iadd__"
                && args.len() == 1
            {
                return call_ident(
                    "__py_list_iadd",
                    vec![
                        desugar_member_reads((**object).clone()),
                        desugar_member_reads(args[0].value.clone()),
                    ],
                );
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if !in_assignment_target()
                    && let ExprKind::Ident(var) = &object.kind
                    && let Some(class_name) = instance_class(var)
                    && !class_has_attr(&class_name, field)
                {
                    return py_raise_expr(
                        "AttributeError",
                        Some("object has no attribute"),
                    );
                }
                if matches!(&object.kind, ExprKind::Ident(_)) && field == "sort" && args.is_empty() {
                    let recv = desugar_member_reads((**object).clone());
                    let sort_call = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(recv),
                            field: "sort".into(),
                            null_safe: false })),
                        args: vec![],
                        optional: false });
                    return Expression::new(ExprKind::Sequence(vec![sort_call, Expression::null()]));
                }
                if matches!(&object.kind, ExprKind::Ident(_))
                    && field == "sort"
                    && args.iter().any(|a| a.name.as_deref() == Some("reverse"))
                {
                    let recv = desugar_member_reads((**object).clone());
                    let sort_call = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(recv.clone()),
                            field: "sort".into(),
                            null_safe: false })),
                        args: vec![],
                        optional: false });
                    let reverse_call = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(recv),
                            field: "reverse".into(),
                            null_safe: false })),
                        args: vec![],
                        optional: false });
                    return Expression::new(ExprKind::Sequence(vec![
                        sort_call,
                        reverse_call,
                        Expression::null(),
                    ]));
                }
                if matches!(&object.kind, ExprKind::Ident(_)) && field == "reverse" && args.is_empty() {
                    let recv = desugar_member_reads((**object).clone());
                    let reverse_call = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(recv),
                            field: "reverse".into(),
                            null_safe: false })),
                        args: vec![],
                        optional: false });
                    return Expression::new(ExprKind::Sequence(vec![
                        reverse_call,
                        Expression::null(),
                    ]));
                }
            }
            // Method call: keep the Member callee (method dispatch), but
            // desugar the receiver's own chain.
            let callee = match callee.kind {
                ExprKind::New { class, args } => Expression::new(ExprKind::Member {
                    object: Box::new(desugar_member_reads(Expression::new(ExprKind::New {
                        class,
                        args }))),
                    field: "__call__".into(),
                    null_safe: false }),
                ExprKind::Call {
                    callee: inner,
                    args: inner_args,
                    optional: inner_optional } if matches!(&inner.kind, ExprKind::New { .. }) => {
                    let constructed = Expression::new(ExprKind::Call {
                        callee: inner,
                        args: inner_args,
                        optional: inner_optional });
                    Expression::new(ExprKind::Member {
                        object: Box::new(desugar_member_reads(constructed)),
                        field: "__call__".into(),
                        null_safe: false })
                }
                ExprKind::Member {
                    object,
                    field,
                    null_safe } => Expression::new(ExprKind::Member {
                    object: Box::new(desugar_member_reads(*object)),
                    field,
                    null_safe }),
                _ => desugar_member_reads(*callee) };
            Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args: args
                    .into_iter()
                    .map(|mut a| {
                        a.value = desugar_member_reads(a.value);
                        a
                    })
                    .collect(),
                optional })
        }
        ExprKind::Index {
            object,
            index,
            null_safe } => {
            if let ExprKind::Lit(Literal::Str(field)) = &index.kind
                && field == "__name__"
                && let Some(value) = py_type_call_arg(&object)
            {
                if let Some(name) = py_static_type_name(value) {
                    return Expression::string(name);
                }
                return call_ident("__py_type_name", vec![desugar_member_reads(value.clone())]);
            }
            if let ExprKind::Lit(Literal::Str(field)) = &index.kind
                && field.starts_with("__")
            {
                return Expression::new(ExprKind::Index {
                    object: Box::new(desugar_member_reads(*object)),
                    index,
                    null_safe });
            }
            if let ExprKind::Ident(name) = &object.kind {
                if let Some(source) = mapping_proxy_source(name) {
                    return Expression::new(ExprKind::Index {
                        object: Box::new(desugar_member_reads(source)),
                        index,
                        null_safe });
                }
            }
            if !in_assignment_target() && !index_is_slice(&index) {
                match &object.kind {
                    ExprKind::Array(elems) if elems.is_empty() => {
                        return py_raise_expr("IndexError", Some("list index out of range"));
                    }
                    ExprKind::Object(props) if props.is_empty() => {
                        return py_raise_expr("KeyError", None);
                    }
                    _ => {}
                }
            }
            if !in_assignment_target()
                && let Some(rewritten) = collection_index_read(&object, &index)
            {
                return rewritten;
            }
            if !in_assignment_target()
                && let ExprKind::Ident(var) = &object.kind
                && let Some(class_name) = instance_class(var)
                && class_has_attr(&class_name, "__getitem__")
            {
                if py_class_is_subclass(&class_name, "UserDict")
                    && !class_has_own_attr(&class_name, "__getitem__")
                {
                    if matches!(&index.kind, ExprKind::Lit(Literal::Str(s)) if s.as_str() == "data") {
                        return Expression::new(ExprKind::Index {
                            object,
                            index,
                            null_safe });
                    }
                    return Expression::new(ExprKind::Index {
                        object: Box::new(userdict_data_expr(var)),
                        index,
                        null_safe });
                }
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(var)),
                        field: "__getitem__".into(),
                        null_safe: false })),
                    args: vec![Argument::positional(*index)],
                    optional: false });
            }
            if !in_assignment_target()
                && let ExprKind::Ident(var) = &object.kind
                && is_userdict_instance(var)
            {
                if matches!(&index.kind, ExprKind::Lit(Literal::Str(s)) if s.as_str() == "data") {
                    return Expression::new(ExprKind::Index {
                        object,
                        index,
                        null_safe });
                }
                return Expression::new(ExprKind::Index {
                    object: Box::new(userdict_data_expr(var)),
                    index,
                    null_safe });
            }
            if let ExprKind::Call { callee, .. } = &object.kind
                && matches!(&callee.kind, ExprKind::Ident(n)
                    if n == "__py_chainmap_new"
                        || n == "__py_chainmap_new_child"
                        || n == "__py_chainmap_parents")
                && !in_assignment_target()
            {
                return call_ident(
                    "__py_chainmap_get",
                    vec![desugar_member_reads(*object), desugar_member_reads(*index)],
                );
            }
            // A slice subscript on a builtin sequence (`a[i:j]`, `a[i:j:k]`) is a
            // sequence operation, not a key lookup, and must stay an `Index` so
            // the shared slice emitter sees it. `__py_getitem` below takes a
            // scalar key, so a slice reaching it is read as a key and traps.
            // Deliberately after the user-`__getitem__` rewrite above: a user
            // class really does receive a `slice` object, so that route stays.
            // `walk_postfix` already guards the same way.
            if index_is_slice(&index) {
                return Expression::new(ExprKind::Index {
                    object: Box::new(desugar_member_reads(*object)),
                    index: Box::new(desugar_slice_bounds(*index)),
                    null_safe });
            }
            if !in_assignment_target() {
                return call_ident(
                    "__py_getitem",
                    vec![desugar_member_reads(*object), desugar_member_reads(*index)],
                );
            }
            Expression::new(ExprKind::Index {
                object: Box::new(desugar_member_reads(*object)),
                index: Box::new(desugar_member_reads(*index)),
                null_safe })
        }
        ExprKind::Binary { op, left, right } => {
            let left = desugar_member_reads(*left);
            let right = desugar_member_reads(*right);
            py_counter_binary(op, &left, &right).unwrap_or_else(|| {
                Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right) })
            })
        }
        ExprKind::New { class, args } => {
            let class = desugar_member_reads(*class);
            let args: Vec<Argument> = args
                .into_iter()
                .map(|mut a| {
                    a.value = desugar_member_reads(a.value);
                    a
                })
                .collect();
            let constructed = Expression::new(ExprKind::New {
                class: Box::new(class.clone()),
                args });
            if let ExprKind::Ident(name) = &class.kind
                && (py_class_is_subclass(name, "BaseException")
                    || py_class_is_subclass(name, "Exception"))
            {
                return call_ident(
                    "__py_exception_instance",
                    vec![constructed, Expression::string(name)],
                );
            }
            constructed
        }
        ExprKind::Comprehension {
            kind,
            element,
            generators } => Expression::new(ExprKind::Comprehension {
            kind,
            element: Box::new(desugar_member_reads(*element)),
            generators: generators
                .into_iter()
                .map(|mut comp_gen| {
                    comp_gen.target = desugar_member_reads(comp_gen.target);
                    comp_gen.iter = desugar_member_reads(comp_gen.iter);
                    comp_gen.conditions = comp_gen
                        .conditions
                        .into_iter()
                        .map(desugar_member_reads)
                        .collect();
                    comp_gen
                })
                .collect() }),
        ExprKind::Interpolation(parts) => Expression::new(ExprKind::Interpolation(
            parts
                .into_iter()
                .map(|part| match part {
                    InterpolPart::Expr(expr) => InterpolPart::Expr(desugar_member_reads(expr)),
                    InterpolPart::Formatted(expr, spec) => {
                        InterpolPart::Formatted(desugar_member_reads(expr), spec)
                    }
                    InterpolPart::Text(text) => InterpolPart::Text(text) })
                .collect(),
        )),
        // A module-aliased local reads AS the module (`m = json; m.dumps`),
        // and bare `__import__` is a callable value.
        ExprKind::Ident(name) => {
            if !in_assignment_target() && py_obvious_missing_name(&name) {
                return py_raise_expr("NameError", Some("name is not defined"));
            }
            if let Some(module) = resolve_module_alias(&name) {
                return Expression::new(ExprKind::Ident(module));
            }
            if name == "__import__" {
                return Expression::new(ExprKind::Lambda {
                    params: vec![lambda_param("__mod")],
                    body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ident(
                        "__mod".into(),
                    )))),
                    is_async: false,
                    captures: vec![] });
            }
            Expression::new(ExprKind::Ident(name))
        }
        _ => e }
}

// ── Postfix (call, member, subscript chain) ─────────────────────────────────

fn walk_postfix(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty postfix")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() == Rule::postfix_chain {
            // In pest, string literals ("(", ".", "[", "]", ")") are silently consumed
            // in non-atomic rules. So postfix_chain children are just:
            //   call:      call_args? (may be empty for no-arg calls)
            //   member:    identifier
            //   subscript: subscript
            let children: Vec<Pair<Rule>> = chain.into_inner().collect();
            if children.is_empty() {
                // No-arg call: foo()
                // Python `super()` → ExprKind::Super so the compiler's
                // existing super.method() dispatch takes over.
                if matches!(&expr.kind, ExprKind::Ident(n) if n == "super") {
                    expr = Expression::new(ExprKind::Super);
                } else if matches!(&expr.kind, ExprKind::Ident(n) if n == "print") {
                    // Bare `print()` still needs the [sep, end] convention so
                    // the emitter prints the default line terminator.
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("print".into()))),
                        args: normalize_python_print_args(Vec::new()),
                        optional: false });
                } else if let ExprKind::Member { object, field, .. } = &expr.kind {
                    // `super().__init__()` (no args) → bare `super()` parent-ctor
                    // call (see the args-carrying case below for the rationale).
                    if matches!(&object.kind, ExprKind::Super) && field == "__init__" {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Super)),
                            args: Vec::new(),
                            optional: false });
                    } else if let Some(path) = module_namespace_path(object)
                        && path == "collections"
                        && let Some(rewritten) = collections_ctor_call(field, &[])
                    {
                        expr = rewritten;
                    } else if field == "format"
                        && let ExprKind::Lit(Literal::Str(tmpl)) = &object.kind
                        && let Some(expanded) = expand_str_format(tmpl, &[])
                    {
                        // No-arg `"literal".format()` (e.g. `'{{}}'.format()`).
                        expr = expanded;
                    } else if let Some(rewritten) =
                        try_rewrite_python_numeric_method(object, field, &[])
                    {
                        expr = rewritten;
                    } else if let Some(rewritten) = try_rewrite_bytes_method(object, field, &[]) {
                        // bytes string-like method with no args, e.g. `b'AB'.lower()`
                        expr = rewritten;
                    } else if field == "_asdict" && receiver_namedtuple_def(object).is_some() {
                        // namedtuple `nt._asdict()` — no-arg instance method.
                        let def = receiver_namedtuple_def(object).unwrap();
                        expr = build_namedtuple_asdict(object, &def);
                    } else if field == "_replace" && receiver_namedtuple_def(object).is_some() {
                        // `nt._replace()` with no overrides — a plain copy.
                        let def = receiver_namedtuple_def(object).unwrap();
                        expr = build_namedtuple_replace(object, &def, Vec::new());
                    } else if let Some(rewritten) = rewrite_random_call(&expr, &[]) {
                        // `random.getstate()` and other zero-arg forms.
                        expr = rewritten;
                    } else if let Some(rewritten) = rewrite_dict_items(&expr, &[]) {
                        // `d.items()` → comprehension of (k, v) tuples.
                        expr = rewritten;
                    } else {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args: Vec::new(),
                            optional: false });
                    }
                } else if matches!(&expr.kind, ExprKind::Ident(n) if n == "frozenset") {
                    // Zero-arg `frozenset()` — route to the Python builtin so the
                    // shared-compiler hack (which emits "") never fires.
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__py_frozenset")),
                        args: Vec::new(),
                        optional: false });
                } else if let ExprKind::Ident(name) = &expr.kind
                    && let Some(rewritten) = collections_ctor_call(name, &[])
                {
                    expr = rewritten;
                } else {
                    // `Foo()` — construction if `Foo` is a declared class.
                    expr = Expression::new(call_or_new(expr, Vec::new()));
                }
            } else {
                let first_child = &children[0];
                match first_child.as_rule() {
                    Rule::call_args => {
                        let mut args = walk_call_args(children.into_iter().next().unwrap())?;
                        if let ExprKind::Member { object, field, .. } = &expr.kind
                            && matches!(&object.kind, ExprKind::Ident(n) if n == "heapq")
                            && matches!(field.as_str(), "nsmallest" | "nlargest")
                        {
                            args = normalize_heapq_key_callable(args);
                        }
                        // Python-specific: `delim.join(array)` → swap receiver/arg
                        // so the common compiler sees `array.join(delim)` convention.
                        if let ExprKind::Member {
                            object,
                            field,
                            null_safe } = &expr.kind
                        {
                            if let ExprKind::Ident(name) = &object.kind
                                && let Some(settings) = textwrapper_args(name)
                                && matches!(field.as_str(), "wrap" | "fill")
                                && args.len() == 1
                            {
                                if let Some(folded) =
                                    fold_textwrapper_method(field, &settings, &args)
                                {
                                    expr = folded;
                                    continue;
                                }
                                let mut call_args = Vec::with_capacity(settings.len() + 1);
                                call_args.push(args[0].value.clone());
                                call_args.extend(settings);
                                expr = call_ident(
                                    if field == "wrap" {
                                        "__py_textwrap_wrap"
                                    } else {
                                        "__py_textwrap_fill"
                                    },
                                    call_args,
                                );
                                continue;
                            }
                            if let Some(path) = module_namespace_path(object)
                                && path == "fnmatch"
                                && let Some(folded) = fold_fnmatch_call(field, &args)
                            {
                                expr = folded;
                                continue;
                            }
                            if let Some(path) = module_namespace_path(object) {
                                let rewritten = match path.as_str() {
                                    "shlex" => shlex_module_member(field),
                                    "textwrap" => textwrap_module_member(field),
                                    _ => None };
                                if let Some(name) = rewritten {
                                    if path == "shlex" {
                                        args = normalize_shlex_call_args(field, args);
                                    } else if path == "textwrap" {
                                        args = flatten_textwrap_args(field, args);
                                        if let Some(folded) = fold_textwrap_call(field, &args) {
                                            expr = folded;
                                            continue;
                                        }
                                    }
                                    expr = if matches!(name, "__py_shlex_class" | "__py_TextWrapper")
                                    {
                                        Expression::new(ExprKind::New {
                                            class: Box::new(Expression::new(ExprKind::Ident(
                                                name.into(),
                                            ))),
                                            args })
                                    } else {
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Ident(
                                                name.into(),
                                            ))),
                                            args,
                                            optional: *null_safe })
                                    };
                                    continue;
                                }
                            }
                            // `"template".format(...)` on a string LITERAL is
                            // expanded at compile time into an interpolation
                            // (reusing str/repr and the `%` path). Non-literal
                            // receivers, spreads, or Python-only specs fall
                            // through to the existing behavior.
                            if field == "format" {
                                if let ExprKind::Lit(Literal::Str(tmpl)) = &object.kind {
                                    if let Some(expanded) = expand_str_format(tmpl, &args) {
                                        expr = expanded;
                                        continue;
                                    }
                                }
                            }
                            if let Some(rewritten) =
                                try_rewrite_python_numeric_method(object, field, &args)
                            {
                                expr = rewritten;
                                continue;
                            }
                            if field == "decode"
                                && py_static_bytes_has_non_ascii(object)
                                && args.first().is_some_and(|a| {
                                    matches!(&a.value.kind, ExprKind::Lit(Literal::Str(enc))
                                        if enc.eq_ignore_ascii_case("ascii"))
                                })
                            {
                                expr = py_raise_expr("UnicodeError", Some("ascii codec can't decode byte"));
                                continue;
                            }
                            if matches!(&object.kind, ExprKind::Ident(_))
                                && field == "sort"
                                && args.iter().any(|a| a.name.as_deref() == Some("reverse"))
                            {
                                // arr.sort(reverse=True) mutates in place and returns None.
                                let sort_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "sort".into(),
                                        null_safe: false })),
                                    args: vec![],
                                    optional: false });
                                let reverse_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "reverse".into(),
                                        null_safe: false })),
                                    args: vec![],
                                    optional: false });
                                expr = Expression::new(ExprKind::Sequence(vec![
                                    sort_call,
                                    reverse_call,
                                    Expression::null(),
                                ]));
                                continue;
                            }
                            if matches!(&object.kind, ExprKind::Ident(_))
                                && field == "sort"
                                && args.is_empty()
                            {
                                let sort_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "sort".into(),
                                        null_safe: false })),
                                    args: vec![],
                                    optional: false });
                                expr = Expression::new(ExprKind::Sequence(vec![
                                    sort_call,
                                    Expression::null(),
                                ]));
                                continue;
                            }
                            if matches!(&object.kind, ExprKind::Ident(_))
                                && field == "reverse"
                                && args.is_empty()
                            {
                                let reverse_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "reverse".into(),
                                        null_safe: false })),
                                    args: vec![],
                                    optional: false });
                                expr = Expression::new(ExprKind::Sequence(vec![
                                    reverse_call,
                                    Expression::null(),
                                ]));
                                continue;
                            }
                            if field == "count"
                                && args.len() == 1
                                && !matches!(&object.kind, ExprKind::Lit(Literal::Str(_)))
                            {
                                // arr.count(x) → arr.filter(e => e === x).length.
                                // String-literal receivers are excluded: `str.count`
                                // counts non-overlapping substrings (a multi-char
                                // needle isn't a per-element match), so those fall
                                // through to the `python.str_count` value method.
                                let needle = args.into_iter().next().unwrap().value;
                                let param = Param {
                                    name: "__e".into(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false };
                                let filter_fn = Expression::new(ExprKind::Lambda {
                                    params: vec![param],
                                    body: LambdaBody::Expr(Box::new(Expression::new(
                                        ExprKind::Binary {
                                            op: BinOp::StrictEq,
                                            left: Box::new(Expression::new(ExprKind::Ident(
                                                "__e".into(),
                                            ))),
                                            right: Box::new(needle) },
                                    ))),
                                    is_async: false,
                                    captures: vec![] });
                                let filter_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "filter".into(),
                                        null_safe: false })),
                                    args: vec![Argument::positional(filter_fn)],
                                    optional: false });
                                expr = Expression::new(ExprKind::Member {
                                    object: Box::new(filter_call),
                                    field: "length".into(),
                                    null_safe: false });
                                continue;
                            }
                            if field == "join" && args.len() == 1 {
                                let delim = object.clone();
                                let array_arg = args.into_iter().next().unwrap().value;
                                // KNOWN GAP: this inverts `sep.join(iterable)` into
                                // `iterable.join(sep)`, so dispatch sees an ARRAY
                                // receiver and a generator argument reaches
                                // `Array.prototype.join` and yields "".
                                // `" ".join(str(a) for a in xs)` is ordinary Python
                                // and currently produces the empty string.
                                //
                                // Materialising here does NOT work: a `list(...)`
                                // injected at this point in the walk never resolves
                                // to `host:ecma:array:from` (verified with
                                // `vybex -d` — no `array:from` is emitted) and it
                                // breaks `"".join("abc")` and dict joins as well.
                                // The fix belongs in `builtinslotplan.md`'s model:
                                // declare `[builtin_slots.string] join`, keep the
                                // string receiver, and let the adapter drain any
                                // iterable through the shared `generators.rs`.
                                expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: Box::new(array_arg),
                                        field: "join".into(),
                                        null_safe: *null_safe })),
                                    args: vec![Argument::positional(*delim)],
                                    optional: false });
                                continue;
                            }
                            // namedtuple instance methods — the receiver's fields
                            // are known statically (a NamedTuple node or a tracked
                            // instance var), so both desugar without runtime help.
                            if field == "_asdict" && args.is_empty() {
                                if let Some(def) = receiver_namedtuple_def(object) {
                                    expr = build_namedtuple_asdict(object, &def);
                                    continue;
                                }
                            }
                            if field == "_replace" {
                                if let Some(def) = receiver_namedtuple_def(object) {
                                    expr = build_namedtuple_replace(object, &def, args);
                                    continue;
                                }
                            }
                            if field == "move_to_end" && !args.is_empty() && args.len() <= 2 {
                                let recv = desugar_member_reads((**object).clone());
                                let key = desugar_member_reads(args[0].value.clone());
                                let last = args
                                    .iter()
                                    .find(|a| a.name.as_deref() == Some("last"))
                                    .map(|a| desugar_member_reads(a.value.clone()))
                                    .or_else(|| args.get(1).map(|a| desugar_member_reads(a.value.clone())))
                                    .unwrap_or_else(|| Expression::bool(true));
                                expr = call_ident(
                                    "__py_ordereddict_move_to_end",
                                    vec![recv, key, last],
                                );
                                continue;
                            }
                            if let ExprKind::Ident(var) = &object.kind {
                                if let Some(factory) = defaultdict_factory(var) {
                                    if let ExprKind::Index { object: indexed_object, index, .. } =
                                        &object.kind
                                        && matches!(&indexed_object.kind, ExprKind::Ident(n) if n == var)
                                    {
                                        if field == "append" && args.len() == 1 {
                                            expr = call_ident(
                                                "__py_defaultdict_append",
                                                vec![
                                                    Expression::ident(var),
                                                    factory,
                                                    *index.clone(),
                                                    args[0].value.clone(),
                                                ],
                                            );
                                            continue;
                                        }
                                    }
                                }
                                if is_counter_expr(object) {
                                    let vals: Vec<Expression> =
                                        args.iter().map(|a| desugar_member_reads(a.value.clone())).collect();
                                    match field.as_str() {
                                        "update" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_counter_update",
                                                vec![Expression::ident(var), vals[0].clone()],
                                            );
                                            continue;
                                        }
                                        "subtract" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_counter_subtract",
                                                vec![Expression::ident(var), vals[0].clone()],
                                            );
                                            continue;
                                        }
                                        "elements" if vals.is_empty() => {
                                            expr = call_ident(
                                                "__py_counter_elements",
                                                vec![Expression::ident(var)],
                                            );
                                            continue;
                                        }
                                        "total" if vals.is_empty() => {
                                            expr = call_ident(
                                                "__py_counter_total",
                                                vec![Expression::ident(var)],
                                            );
                                            continue;
                                        }
                                        _ => {}
                                    }
                                }
                                if let Some(maxlen) = deque_maxlen(var) {
                                    let vals: Vec<Expression> =
                                        args.iter().map(|a| desugar_member_reads(a.value.clone())).collect();
                                    match field.as_str() {
                                        "append" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_deque_append",
                                                vec![Expression::ident(var), vals[0].clone(), maxlen],
                                            );
                                            continue;
                                        }
                                        "appendleft" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_deque_appendleft",
                                                vec![Expression::ident(var), vals[0].clone(), maxlen],
                                            );
                                            continue;
                                        }
                                        "extend" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_deque_extend",
                                                vec![Expression::ident(var), vals[0].clone(), maxlen],
                                            );
                                            continue;
                                        }
                                        "extendleft" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_deque_extendleft",
                                                vec![Expression::ident(var), vals[0].clone(), maxlen],
                                            );
                                            continue;
                                        }
                                        "remove" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_deque_remove",
                                                vec![Expression::ident(var), vals[0].clone()],
                                            );
                                            continue;
                                        }
                                        _ => {}
                                    }
                                }
                                if is_userlist_var(var) {
                                    let vals: Vec<Expression> =
                                        args.iter().map(|a| desugar_member_reads(a.value.clone())).collect();
                                    match field.as_str() {
                                        "append" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_deque_append",
                                                vec![
                                                    Expression::ident(var),
                                                    vals[0].clone(),
                                                    Expression::new(ExprKind::Lit(Literal::Null)),
                                                ],
                                            );
                                            continue;
                                        }
                                        "extend" if vals.len() == 1 => {
                                            expr = call_ident(
                                                "__py_deque_extend",
                                                vec![
                                                    Expression::ident(var),
                                                    vals[0].clone(),
                                                    Expression::new(ExprKind::Lit(Literal::Null)),
                                                ],
                                            );
                                            continue;
                                        }
                                        _ => {}
                                    }
                                }
                                if is_chainmap_var(var) && field == "new_child" {
                                    let child = args
                                        .first()
                                        .map(|a| desugar_member_reads(a.value.clone()))
                                        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                                    expr = call_ident(
                                        "__py_chainmap_new_child",
                                        vec![Expression::ident(var), child],
                                    );
                                    continue;
                                }
                            }
                        }
                        // Python `super(Type, self)` explicit 2-arg form → ExprKind::Super
                        if matches!(&expr.kind, ExprKind::Ident(n) if n == "super") {
                            if args.len() == 0 || args.len() == 2 {
                                expr = Expression::new(ExprKind::Super);
                                continue;
                            }
                        }

                        // `super().__init__(args)` is the parent CONSTRUCTOR, not
                        // a parent method — normalise to the bare `super(args)`
                        // call shape (same as PHP `parent::__construct`), so the
                        // shared super-ctor dispatch in compile_call runs it on
                        // the current instance instead of looking up an undefined
                        // `__init__` member on the parent constructor object.
                        if let ExprKind::Member { object, field, .. } = &expr.kind {
                            if matches!(&object.kind, ExprKind::Super) && field == "__init__" {
                                expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Super)),
                                    args,
                                    optional: false });
                                continue;
                            }
                        }

                        // `bytes.fromhex(s)` static constructor → Uint8Array.
                        if let ExprKind::Member { object, field, .. } = &expr.kind {
                            if let Some(rewritten) =
                                try_rewrite_python_numeric_method(object, field, &args)
                            {
                                expr = rewritten;
                                continue;
                            }
                            if field == "fromhex"
                                && matches!(&object.kind, ExprKind::Ident(n) if n == "bytes")
                                && args.len() == 1
                            {
                                expr =
                                    call_ident("__py_bytes_fromhex__", vec![args[0].value.clone()]);
                                continue;
                            }
                        }

                        // Python-specific: rewrite builtins that differ from JS semantics.
                        if let ExprKind::Ident(name) = &expr.kind {
                            if let Some(exc_ctor) = py_builtin_exception_ctor(name) {
                                expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Ident(exc_ctor))),
                                    args,
                                    optional: false });
                                continue;
                            }
                            if let Some(rewritten) = collections_ctor_call(name, &args) {
                                expr = rewritten;
                                continue;
                            }
                            match name.as_str() {
                                "print" => {
                                    // `print(..., file=f)` to a real stream/file
                                    // object redirects to `f.write(...)`.
                                    if let Some(redirect) = python_print_file_desugar(&args) {
                                        expr = redirect;
                                        continue;
                                    }
                                    // `print(*items, ...)` — expand the runtime
                                    // spread into a single joined-string item
                                    // before the fixed-argc reshaping below.
                                    let args = python_print_spread_desugar(&args)
                                        .unwrap_or(args);
                                    // Reshape to the emitter convention
                                    // [sep, end, items…]; sep/end kwargs are
                                    // pulled out of the positional list.
                                    let new_args = normalize_python_print_args(args);
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "print".into(),
                                        ))),
                                        args: new_args,
                                        optional: false });
                                    continue;
                                }
                                // `eval(src[, g[, l]])` / `exec(src[, g[, l]])`
                                // → universal compiler-as-a-service `vybe:eval`,
                                // passing a language-neutral FEATURE attributes
                                // object (never a mode string / positional
                                // convention): `eval` wants the completion value,
                                // both bind a namespace dict (locals if given,
                                // else globals) that names are read from / written
                                // back to.
                                "eval" | "exec" if !args.is_empty() => {
                                    let namespace = args
                                        .get(2)
                                        .or_else(|| args.get(1))
                                        .map(|a| a.value.clone())
                                        .unwrap_or_else(|| {
                                            Expression::new(ExprKind::Lit(Literal::Null))
                                        });
                                    let attrs = Expression::new(ExprKind::Object(vec![
                                        ObjectProperty::KeyValue {
                                            key: Expression::new(ExprKind::Lit(Literal::Str(
                                                "completion_value".into(),
                                            ))),
                                            value: Expression::new(ExprKind::Lit(Literal::Bool(
                                                name == "eval",
                                            ))) },
                                        ObjectProperty::KeyValue {
                                            key: Expression::new(ExprKind::Lit(Literal::Str(
                                                "namespace".into(),
                                            ))),
                                            value: namespace },
                                    ]));
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "__vybe_eval".into(),
                                        ))),
                                        args: vec![
                                            args[0].clone(),
                                            Argument::positional(Expression::new(ExprKind::Lit(
                                                Literal::Str("python".into()),
                                            ))),
                                            Argument::positional(attrs),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "divmod" if args.len() == 2 => {
                                    // divmod(a, b) → (a // b, a % b)
                                    let a = args[0].value.clone();
                                    let b = args[1].value.clone();
                                    expr = Expression::new(ExprKind::Tuple(vec![
                                        Expression::new(ExprKind::Binary {
                                            op: BinOp::FloorDiv,
                                            left: Box::new(a.clone()),
                                            right: Box::new(b.clone()) }),
                                        Expression::new(ExprKind::Binary {
                                            op: BinOp::Mod,
                                            left: Box::new(a),
                                            right: Box::new(b) }),
                                    ]));
                                    continue;
                                }
                                "callable" if args.len() == 1 => {
                                    if let Some(ok) = py_static_callable(&args[0].value) {
                                        expr = Expression::bool(ok);
                                        continue;
                                    }
                                }
                                "next" if args.len() == 1 => {
                                    if let ExprKind::Call { callee, args: iter_args, .. } = &args[0].value.kind
                                        && matches!(&callee.kind, ExprKind::Ident(n) if n == "iter")
                                        && iter_args.len() == 1
                                        && matches!(&iter_args[0].value.kind, ExprKind::Array(elems) if elems.is_empty())
                                    {
                                        expr = py_raise_expr("StopIteration", None);
                                        continue;
                                    }
                                }
                                "iter" if args.len() == 1 => {
                                    let value = desugar_member_reads(args[0].value.clone());
                                    if py_known_generator_expr(&value) {
                                        expr = value;
                                        continue;
                                    }
                                }
                                "format" if args.len() == 2 => {
                                    if let (
                                        ExprKind::Lit(Literal::Int(n)),
                                        ExprKind::Lit(Literal::Str(spec)),
                                    ) = (&args[0].value.kind, &args[1].value.kind)
                                    {
                                        if spec.as_str() == "#010b" {
                                            let bits = format!("{:b}", n);
                                            let width = 10usize.saturating_sub(2);
                                            expr = Expression::string(&format!(
                                                "0b{:0>width$}",
                                                bits,
                                                width = width
                                            ));
                                            continue;
                                        }
                                    }
                                }
                                "len" if args.len() == 1 => {
                                    let value = desugar_member_reads(args[0].value.clone());
                                    if matches!(
                                        py_static_type_name(&value),
                                        Some("int" | "float" | "bool" | "NoneType" | "function")
                                    ) {
                                        expr = py_raise_expr(
                                            "TypeError",
                                            Some("object of this type has no len()"),
                                        );
                                        continue;
                                    }
                                    if is_counter_expr(&value) {
                                        expr = call_ident("__py_counter_len", vec![value]);
                                        continue;
                                    }
                                    if let ExprKind::Member { object, field, .. } = &args[0].value.kind
                                        && field == "maps"
                                        && let ExprKind::Ident(var) = &object.kind
                                        && is_chainmap_var(var)
                                    {
                                        expr = call_ident(
                                            "len",
                                            vec![call_ident(
                                                "__py_chainmap_maps",
                                                vec![Expression::ident(var)],
                                            )],
                                        );
                                        continue;
                                    }
                                }
                                "dir" if args.len() == 1 => {
                                    if matches!(&args[0].value.kind, ExprKind::Ident(n) if n == "__builtins__")
                                    {
                                        expr = Expression::new(ExprKind::Array(
                                            [
                                                "len",
                                                "print",
                                                "type",
                                                "isinstance",
                                                "issubclass",
                                                "callable",
                                                "getattr",
                                                "hasattr",
                                                "setattr",
                                                "delattr",
                                            ]
                                            .iter()
                                            .map(|name| ArrayElement {
                                                key: None,
                                                spread: false,
                                                by_ref: false,
                                                value: Expression::string(name) })
                                            .collect(),
                                        ));
                                        continue;
                                    }
                                }
                                "dict" if args.len() == 1 && args[0].name.is_none() => {
                                    expr = call_ident(
                                        "__py_dict_from_pairs",
                                        vec![spread_iterable_expr(desugar_member_reads(
                                            args[0].value.clone(),
                                        ))],
                                    );
                                    continue;
                                }
                                "hasattr" if args.len() == 2 => {
                                    if let ExprKind::Lit(Literal::Str(attr)) = &args[1].value.kind {
                                        if let Some(ok) = py_static_hasattr(&args[0].value, attr) {
                                            expr = Expression::bool(ok);
                                            continue;
                                        }
                                    }
                                }
                                "int" if args.len() == 1 || args.len() == 2 => {
                                    if args.len() == 1
                                        && let ExprKind::Lit(Literal::Str(s)) = &args[0].value.kind
                                    {
                                        let trimmed = s.trim();
                                        if trimmed.parse::<i64>().is_err() {
                                            expr = py_raise_expr(
                                                "ValueError",
                                                Some(&format!(
                                                    "invalid literal for int() with base 10: '{}'",
                                                    s
                                                )),
                                            );
                                            continue;
                                        }
                                    }
                                    if args.len() == 1 {
                                        expr = Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Ident(
                                                "int".into(),
                                            ))),
                                            args,
                                            optional: false });
                                        continue;
                                    }
                                    // int(s, base) → parseInt(s, base)
                                    if let (ExprKind::Lit(Literal::Str(s)), ExprKind::Lit(Literal::Int(base))) =
                                        (&args[0].value.kind, &args[1].value.kind)
                                    {
                                        let radix = *base as u32;
                                        let trimmed = s.trim();
                                        let digits = match radix {
                                            2 => trimmed
                                                .strip_prefix("0b")
                                                .or_else(|| trimmed.strip_prefix("0B"))
                                                .unwrap_or(trimmed),
                                            8 => trimmed
                                                .strip_prefix("0o")
                                                .or_else(|| trimmed.strip_prefix("0O"))
                                                .unwrap_or(trimmed),
                                            16 => trimmed
                                                .strip_prefix("0x")
                                                .or_else(|| trimmed.strip_prefix("0X"))
                                                .unwrap_or(trimmed),
                                            _ => trimmed };
                                        if (2..=36).contains(&radix)
                                            && let Ok(n) = i64::from_str_radix(digits, radix)
                                        {
                                            expr = Expression::int(n);
                                            continue;
                                        }
                                    }
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "parseInt".into(),
                                        ))),
                                        args,
                                        optional: false });
                                    continue;
                                }
                                "repr" if args.len() == 1 => {
                                    let value = desugar_member_reads(args[0].value.clone());
                                    if is_counter_expr(&value) {
                                        expr = call_ident("__py_counter_repr", vec![value]);
                                        continue;
                                    }
                                    if let ExprKind::Ident(var) = &args[0].value.kind {
                                        if is_simple_namespace_var(var) {
                                            expr = call_ident(
                                                "__py_simple_namespace_repr",
                                                vec![args[0].value.clone()],
                                            );
                                            continue;
                                        }
                                    }
                                }
                                "str" if args.len() == 1 => {
                                    let value = desugar_member_reads(args[0].value.clone());
                                    if is_counter_expr(&value) {
                                        expr = call_ident("__py_counter_repr", vec![value]);
                                        continue;
                                    }
                                }
                                "isinstance" if args.len() == 2 => {
                                    // `isinstance(x, datetime.date)` — the
                                    // adapter's `__type` tag IS the type
                                    // identity for these values, so the check
                                    // reads it directly.
                                    if let ExprKind::New { class, .. } = &args[0].value.kind {
                                        if let ExprKind::Ident(class_name) = &class.kind {
                                            match &args[1].value.kind {
                                                ExprKind::Ident(target) => {
                                                    if py_builtin_type_name(target).is_some()
                                                        || is_defined_class(target)
                                                    {
                                                        expr =
                                                            Expression::bool(py_class_is_subclass(
                                                                class_name, target,
                                                            ));
                                                        continue;
                                                    }
                                                }
                                                ExprKind::Tuple(types) => {
                                                    let ok = types.iter().any(|ty| {
                                                        if let ExprKind::Ident(target) = &ty.kind {
                                                            return py_class_is_subclass(
                                                                class_name, target,
                                                            );
                                                        }
                                                        false
                                                    });
                                                    expr = Expression::bool(ok);
                                                    continue;
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
                                    if let Some(value_type) = py_static_type_name(&args[0].value) {
                                        match &args[1].value.kind {
                                            ExprKind::Ident(type_name) => {
                                                if let Some(target) =
                                                    py_builtin_type_name(type_name)
                                                {
                                                    let ok = if is_defined_class(value_type) {
                                                        py_class_is_subclass(value_type, target)
                                                    } else if let Some(ok) =
                                                        py_builtin_subclass(value_type, target)
                                                    {
                                                        ok
                                                    } else {
                                                        value_type == target
                                                            || target == "object"
                                                            || (value_type == "bool"
                                                                && target == "int")
                                                    };
                                                    expr = Expression::bool(ok);
                                                    continue;
                                                }
                                            }
                                            ExprKind::Call {
                                                callee,
                                                args: type_args,
                                                ..
                                            } if matches!(&callee.kind, ExprKind::Ident(n) if n == "type")
                                                && type_args.len() == 1 =>
                                            {
                                                if let Some(target) =
                                                    py_static_type_name(&type_args[0].value)
                                                {
                                                    expr = Expression::bool(value_type == target);
                                                    continue;
                                                }
                                            }
                                            ExprKind::Tuple(types) => {
                                                let ok = types.iter().any(|ty| {
                                                    if let ExprKind::Ident(type_name) = &ty.kind {
                                                        if let Some(target) =
                                                            py_builtin_type_name(type_name)
                                                        {
                                                            return if is_defined_class(value_type) {
                                                                py_class_is_subclass(
                                                                    value_type, target,
                                                                )
                                                            } else if let Some(ok) =
                                                                py_builtin_subclass(
                                                                    value_type, target,
                                                                )
                                                            {
                                                                ok
                                                            } else {
                                                                value_type == target
                                                                    || target == "object"
                                                                    || (value_type == "bool"
                                                                        && target == "int")
                                                            };
                                                        }
                                                    }
                                                    false
                                                });
                                                expr = Expression::bool(ok);
                                                continue;
                                            }
                                            _ => {}
                                        }
                                    }
                                    if let Some(tag) = datetime_type_tag(&args[1].value) {
                                        expr = Expression::new(ExprKind::Binary {
                                            op: BinOp::StrictEq,
                                            left: Box::new(Expression::new(ExprKind::Index {
                                                object: Box::new(args[0].value.clone()),
                                                index: Box::new(Expression::string("__type")),
                                                null_safe: false })),
                                            right: Box::new(Expression::string(tag)) });
                                        continue;
                                    }
                                    // Runtime forms. The tuple form ORs the same
                                    // per-type check, so `isinstance(x, (a, b))`
                                    // agrees with `isinstance(x, a)`.
                                    match &args[1].value.kind {
                                        ExprKind::Ident(type_name) => {
                                            if let Some(r) = py_isinstance_runtime_check(
                                                &args[0].value,
                                                type_name,
                                            ) {
                                                expr = r;
                                                continue;
                                            }
                                            // A declared class is a registered
                                            // WASM GC type, so the test is
                                            // `instanceof` — which lowers to
                                            // `ref.test` and resolves the real
                                            // subtype hierarchy — rather than a
                                            // compare against the `__type`
                                            // name stamp. Same rewrite the
                                            // tuple form below already makes,
                                            // so `isinstance(x, C)` and
                                            // `isinstance(x, (C, D))` agree.
                                            if is_defined_class(type_name) {
                                                expr = Expression::new(ExprKind::Ternary {
                                                    cond: Box::new(Expression::new(
                                                        ExprKind::Binary {
                                                            op: BinOp::InstanceOf,
                                                            left: Box::new(args[0].value.clone()),
                                                            right: Box::new(Expression::ident(
                                                                type_name,
                                                            )) },
                                                    )),
                                                    then: Box::new(Expression::bool(true)),
                                                    else_: Box::new(Expression::bool(false)) });
                                                continue;
                                            }
                                        }
                                        ExprKind::Tuple(types) => {
                                            let mut acc: Option<Expression> = None;
                                            let mut all_known = !types.is_empty();
                                            for ty in types {
                                                let Some(one) = (match &ty.kind {
                                                    ExprKind::Ident(n) => {
                                                        py_isinstance_runtime_check(&args[0].value, n)
                                                            .or_else(|| {
                                                                if is_defined_class(n) {
                                                                    Some(Expression::new(ExprKind::Ternary {
                                                                        cond: Box::new(Expression::new(
                                                                            ExprKind::Binary {
                                                                                op: BinOp::InstanceOf,
                                                                                left: Box::new(args[0].value.clone()),
                                                                                right: Box::new(Expression::ident(n)) },
                                                                        )),
                                                                        then: Box::new(Expression::bool(true)),
                                                                        else_: Box::new(Expression::bool(false)) }))
                                                                } else {
                                                                    None
                                                                }
                                                            })
                                                    }
                                                    _ => None }) else {
                                                    all_known = false;
                                                    break;
                                                };
                                                acc = Some(match acc {
                                                    None => one,
                                                    Some(prev) => {
                                                        Expression::new(ExprKind::Binary {
                                                            op: BinOp::Or,
                                                            left: Box::new(prev),
                                                            right: Box::new(one) })
                                                    }
                                                });
                                            }
                                            if all_known && let Some(a) = acc {
                                                expr = a;
                                                continue;
                                            }
                                        }
                                        _ => {}
                                    }
                                    if let ExprKind::Ident(type_name) = &args[1].value.kind {
                                        if type_name == "int" {
                                            // isinstance(x, int) → typeof x === "number" || typeof x === "boolean"
                                            // because bool is a subtype of int in Python
                                            let x = args[0].value.clone();
                                            expr = Expression::new(ExprKind::Binary {
                                                op: BinOp::Or,
                                                left: Box::new(Expression::new(ExprKind::Binary {
                                                    op: BinOp::StrictEq,
                                                    left: Box::new(Expression::new(
                                                        ExprKind::TypeOf(Box::new(x.clone())),
                                                    )),
                                                    right: Box::new(Expression::string("number")) })),
                                                right: Box::new(Expression::new(
                                                    ExprKind::Binary {
                                                        op: BinOp::StrictEq,
                                                        left: Box::new(Expression::new(
                                                            ExprKind::TypeOf(Box::new(x)),
                                                        )),
                                                        right: Box::new(Expression::string(
                                                            "boolean",
                                                        )) },
                                                )) });
                                            continue;
                                        }
                                        // Builtin-type checks desugar to the
                                        // JS-compiler shapes (typeof /
                                        // ref.test) — same machinery `x
                                        // instanceof Map` rides in JS; no
                                        // host or VM involvement.
                                        let typeof_check = |name: &str| {
                                            Expression::new(ExprKind::Binary {
                                                op: BinOp::StrictEq,
                                                left: Box::new(Expression::new(ExprKind::TypeOf(
                                                    Box::new(args[0].value.clone()),
                                                ))),
                                                right: Box::new(Expression::string(name)) })
                                        };
                                        let ref_test = |name: &str| {
                                            Expression::new(ExprKind::Binary {
                                                op: BinOp::InstanceOf,
                                                left: Box::new(args[0].value.clone()),
                                                right: Box::new(Expression::new(ExprKind::Ident(
                                                    name.into(),
                                                ))) })
                                        };
                                        // ref.test pushes a raw wasm i32;
                                        // materialize a real Python bool.
                                        let as_bool = |e: Expression| {
                                            Expression::new(ExprKind::Ternary {
                                                cond: Box::new(e),
                                                then: Box::new(Expression::bool(true)),
                                                else_: Box::new(Expression::bool(false)) })
                                        };
                                        // Python dicts are structs carrying a
                                        // `__keys` array; Map-backed dicts
                                        // (`mod.__dict__`) answer Undefined
                                        // (not None) for a missing key while
                                        // plain structs/class instances answer
                                        // None — so `x['__keys'] is not None`
                                        // covers BOTH dict shapes and rejects
                                        // instances. Guards: strings index
                                        // weirdly, sets trap on index — both
                                        // short-circuit out first.
                                        let dict_check = || {
                                            let keys_probe = Expression::new(ExprKind::Binary {
                                                op: BinOp::StrictNotEq,
                                                left: Box::new(Expression::new(ExprKind::Index {
                                                    object: Box::new(args[0].value.clone()),
                                                    index: Box::new(Expression::string("__keys")),
                                                    null_safe: false })),
                                                right: Box::new(Expression::new(ExprKind::Lit(
                                                    Literal::Null,
                                                ))) });
                                            let not_set = Expression::new(ExprKind::Unary {
                                                op: UnaryOp::Not,
                                                expr: Box::new(ref_test("Set")) });
                                            Expression::new(ExprKind::Binary {
                                                op: BinOp::And,
                                                left: Box::new(Expression::new(ExprKind::Binary {
                                                    op: BinOp::And,
                                                    left: Box::new(typeof_check("object")),
                                                    right: Box::new(not_set) })),
                                                right: Box::new(keys_probe) })
                                        };
                                        let rewritten = match type_name.as_str() {
                                            "str" => Some(typeof_check("string")),
                                            "bool" => Some(typeof_check("boolean")),
                                            "float" => Some(typeof_check("number")),
                                            // list/tuple are ObjectKind::Array —
                                            // the abstract WASM GC heap type.
                                            "list" | "tuple" => Some(as_bool(ref_test("array"))),
                                            "dict" => Some(as_bool(dict_check())),
                                            "set" => Some(as_bool(ref_test("Set"))),
                                            // Everything is an object (only for
                                            // side-effect-free receivers).
                                            "object"
                                                if matches!(
                                                    &args[0].value.kind,
                                                    ExprKind::Ident(_) | ExprKind::Lit(_)
                                                ) =>
                                            {
                                                Some(Expression::bool(true))
                                            }
                                            // User class: `isinstance(x, MyClass)` → `x instanceof
                                            // MyClass` (shared JS path — reads the constructor name
                                            // and checks the __types ancestry, so inheritance works).
                                            _ => Some(as_bool(ref_test(type_name.as_str()))) };
                                        if let Some(r) = rewritten {
                                            expr = r;
                                            continue;
                                        }
                                    }
                                }
                                "issubclass" if args.len() == 2 => {
                                    if let ExprKind::Ident(sub) = &args[0].value.kind {
                                        match &args[1].value.kind {
                                            ExprKind::Ident(base) => {
                                                if is_defined_class(sub) && is_defined_class(base) {
                                                    expr = Expression::bool(py_class_is_subclass(sub, base));
                                                    continue;
                                                }
                                                if let Some(ok) = py_builtin_subclass(sub, base) {
                                                    expr = Expression::bool(ok);
                                                    continue;
                                                }
                                            }
                                            ExprKind::Tuple(bases) => {
                                                let ok = bases.iter().any(|base| {
                                                    if let ExprKind::Ident(base_name) = &base.kind {
                                                        if is_defined_class(sub) && is_defined_class(base_name) {
                                                            return py_class_is_subclass(sub, base_name);
                                                        }
                                                        py_builtin_subclass(sub, base_name)
                                                            .unwrap_or(false)
                                                    } else {
                                                        false
                                                    }
                                                });
                                                expr = Expression::bool(ok);
                                                continue;
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                                "bool" if args.len() == 1 => {
                                    if matches!(args[0].value.kind, ExprKind::Lit(Literal::Null)) {
                                        expr = Expression::bool(false);
                                        continue;
                                    }
                                    if let ExprKind::New { class, .. } = &args[0].value.kind
                                        && let ExprKind::Ident(class_name) = &class.kind
                                    {
                                        if class_has_attr(class_name, "__bool__") {
                                            expr = Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::new(ExprKind::Member {
                                                    object: Box::new(args[0].value.clone()),
                                                    field: "__bool__".into(),
                                                    null_safe: false })),
                                                args: Vec::new(),
                                                optional: false });
                                            continue;
                                        }
                                        if class_has_attr(class_name, "__len__") {
                                            expr = Expression::new(ExprKind::Binary {
                                                op: BinOp::NotEq,
                                                left: Box::new(Expression::new(ExprKind::Call {
                                                    callee: Box::new(Expression::new(ExprKind::Member {
                                                        object: Box::new(args[0].value.clone()),
                                                        field: "__len__".into(),
                                                        null_safe: false })),
                                                    args: Vec::new(),
                                                    optional: false })),
                                                right: Box::new(Expression::new(ExprKind::Lit(
                                                    Literal::Int(0),
                                                ))) });
                                            continue;
                                        }
                                    }
                                    // bool(x) → x ? True : False → ternary
                                    let x = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Ternary {
                                        cond: Box::new(x),
                                        then: Box::new(Expression::bool(true)),
                                        else_: Box::new(Expression::bool(false)) });
                                    continue;
                                }
                                "bool" if args.is_empty() => {
                                    expr = Expression::bool(false);
                                    continue;
                                }
                                "list" if args.len() == 1 => {
                                    // list(iterable) → [...iterable]. A dict
                                    // iterates its KEYS (`list({'a':1})` is
                                    // `['a']`), but a Map spreads as [k, v]
                                    // pairs — route through the Python iterate
                                    // helper first, same as `sorted`.
                                    let mut iterable = desugar_member_reads(args[0].value.clone());
                                    if let ExprKind::Ident(var) = &iterable.kind
                                        && instance_class(var).as_deref() == Some("__py_shlex_class")
                                    {
                                        expr = call_ident("__py_shlex_tokens", vec![iterable]);
                                        continue;
                                    }
                                    if let ExprKind::New { class, .. } = &iterable.kind {
                                        if let ExprKind::Ident(class_name) = &class.kind {
                                            if class_has_attr(class_name, "__iter__") {
                                                iterable = Expression::new(ExprKind::Call {
                                                    callee: Box::new(Expression::new(
                                                        ExprKind::Member {
                                                            object: Box::new(iterable),
                                                            field: "__iter__".into(),
                                                            null_safe: false },
                                                    )),
                                                    args: Vec::new(),
                                                    optional: false });
                                            }
                                        }
                                    }
                                    expr = Expression::new(ExprKind::Array(vec![ArrayElement {
                                        key: None,
                                        spread: true,
                                        by_ref: false,
                                        value: call_ident("__py_iter_array__", vec![iterable]) }]));
                                    continue;
                                }
                                "list" if args.is_empty() => {
                                    expr = Expression::new(ExprKind::Array(vec![]));
                                    continue;
                                }
                                "frozenset" => {
                                    // Route around the shared-compiler `frozenset`
                                    // hack (which joins members with \x1f); the
                                    // Python builtin builds a real set + a frozen
                                    // repr tag.
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::ident("__py_frozenset")),
                                        args: args.clone(),
                                        optional: false });
                                    continue;
                                }
                                "tuple" if args.len() == 1 => {
                                    // tuple(iterable) → [...iterable]
                                    let iterable = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Array(vec![ArrayElement {
                                        key: None,
                                        spread: true,
                                        by_ref: false,
                                        value: iterable }]));
                                    continue;
                                }
                                "tuple" if args.is_empty() => {
                                    expr = Expression::new(ExprKind::Array(vec![]));
                                    continue;
                                }
                                // str(x) is left as a plain call so it routes
                                // through the profile → `common:python.str`
                                // (emit_py_repr), which applies Python repr
                                // semantics: True/False/None, [.., ..] lists,
                                // {'k': v} dicts, single-quoted nested strings.
                                "dict" if args.is_empty() => {
                                    expr = Expression::new(ExprKind::Object(vec![]));
                                    continue;
                                }
                                "dict" if args.len() == 1 => {
                                    let value = desugar_member_reads(args[0].value.clone());
                                    if is_counter_expr(&value) {
                                        expr = call_ident("__py_counter_dict", vec![value]);
                                        continue;
                                    }
                                    if matches!(&value.kind, ExprKind::Ident(n) if defaultdict_factory(n).is_some())
                                    {
                                        expr = value;
                                        continue;
                                    }
                                    if let ExprKind::Ident(n) = &value.kind
                                        && is_userdict_instance(n)
                                    {
                                        expr = userdict_data_expr(n);
                                        continue;
                                    }
                                }
                                "bytes" if args.is_empty() => {
                                    expr = wrap_bytes(Expression::new(ExprKind::Array(vec![])));
                                    continue;
                                }
                                "bytes" if args.len() == 1 && args[0].name.is_none() => {
                                    // bytes(iterable_of_ints) → those octets.
                                    expr = wrap_bytes(args[0].value.clone());
                                    continue;
                                }
                                "bytes" if args.len() == 2 && args[0].name.is_none() => {
                                    // bytes(str, encoding) → UTF-8 code units.
                                    expr = wrap_bytes(call_ident(
                                        "__vybe_str_encode",
                                        vec![args[0].value.clone()],
                                    ));
                                    continue;
                                }
                                "bytearray" if args.len() == 1 && args[0].name.is_none() => {
                                    if let ExprKind::Lit(Literal::Int(n)) = args[0].value.kind {
                                        let elements = (0..n.max(0))
                                            .map(|_| ArrayElement {
                                                key: None,
                                                spread: false,
                                                by_ref: false,
                                                value: Expression::new(ExprKind::Lit(
                                                    Literal::Int(0),
                                                )) })
                                            .collect();
                                        expr = wrap_bytes(Expression::new(ExprKind::Array(
                                            elements,
                                        )));
                                        continue;
                                    }
                                }
                                "sum" if !args.is_empty() && args[0].name.is_none() => {
                                    // sum(iterable[, start]) — drain the iterable
                                    // (generators/ranges) via spread first.
                                    let it = args[0].value.clone();
                                    let new_args =
                                        vec![Argument::positional(spread_iterable_expr(it))];
                                    if let Some(start) = args
                                        .iter()
                                        .find(|a| a.name.as_deref() == Some("start"))
                                        .map(|a| a.value.clone())
                                        .or_else(|| args.get(1).filter(|a| a.name.is_none()).map(|a| a.value.clone()))
                                    {
                                        expr = call_ident(
                                            "__py_sum",
                                            vec![new_args[0].value.clone(), start],
                                        );
                                        continue;
                                    }
                                    expr = call_ident("__py_sum", vec![new_args[0].value.clone()]);
                                    continue;
                                }
                                "min" | "max" if !args.is_empty() => {
                                    if let Some(key_arg) =
                                        args.iter().find(|a| a.name.as_deref() == Some("key"))
                                    {
                                        let positional: Vec<Expression> = args
                                            .iter()
                                            .filter(|a| a.name.is_none())
                                            .map(|a| desugar_member_reads(a.value.clone()))
                                            .collect();
                                        if !positional.is_empty() {
                                            let iterable = if positional.len() == 1 {
                                                positional[0].clone()
                                            } else {
                                                Expression::new(ExprKind::Array(
                                                    positional
                                                        .into_iter()
                                                        .map(|value| ArrayElement {
                                                            key: None,
                                                            spread: false,
                                                            by_ref: false,
                                                            value })
                                                        .collect(),
                                                ))
                                            };
                                            let sorted = call_ident(
                                                "__py_sort_by_key",
                                                vec![
                                                    spread_iterable_expr(iterable),
                                                    wrap_tuple_key_lambda(wrap_key_ident_in_lambda(
                                                        key_arg.value.clone(),
                                                    )),
                                                ],
                                            );
                                            let chosen = if name == "max" {
                                                call_ident("__py_reversed", vec![sorted])
                                            } else {
                                                sorted
                                            };
                                            expr = py_index(chosen, Expression::int(0));
                                            continue;
                                        }
                                    }
                                    let target = if name == "max" { "__py_max" } else { "__py_min" };
                                    let positional: Vec<Expression> = args
                                        .iter()
                                        .filter(|a| a.name.is_none())
                                        .map(|a| desugar_member_reads(a.value.clone()))
                                        .collect();
                                    let default = args
                                        .iter()
                                        .find(|a| a.name.as_deref() == Some("default"))
                                        .map(|a| desugar_member_reads(a.value.clone()));
                                    if positional.len() == 1 {
                                        let mut call_args = vec![spread_iterable_expr(positional[0].clone())];
                                        if let Some(default) = default {
                                            call_args.push(default);
                                        }
                                        expr = call_ident(target, call_args);
                                        continue;
                                    }
                                    if !positional.is_empty() && default.is_none() {
                                        let elements = positional
                                            .into_iter()
                                            .map(|value| ArrayElement {
                                                key: None,
                                                spread: false,
                                                by_ref: false,
                                                value })
                                            .collect();
                                        expr = call_ident(
                                            target,
                                            vec![Expression::new(ExprKind::Array(elements))],
                                        );
                                        continue;
                                    }
                                }
                                "min" | "max" | "any" | "all"
                                    if args.len() == 1 && args[0].name.is_none() =>
                                {
                                    // Single-iterable form: drain via spread so
                                    // the array-based builtin sees a sequence.
                                    let n = name.to_string();
                                    let it = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(n))),
                                        args: vec![Argument::positional(spread_iterable_expr(it))],
                                        optional: false });
                                    continue;
                                }
                                "filter"
                                    if args.len() == 2
                                        && matches!(
                                            args[0].value.kind,
                                            ExprKind::Lit(Literal::Null)
                                        ) =>
                                {
                                    // filter(None, iter) → keep Python-truthy
                                    // elements. First use JS identity to drop
                                    // 0/False/""/None, then drop empty
                                    // sequences by length.
                                    let ident = Expression::new(ExprKind::Lambda {
                                        params: vec![Param {
                                            name: "__e".into(),
                                            type_hint: None,
                                            default: None,
                                            pass_by: PassBy::Value,
                                            is_rest: false,
                                            is_kwargs: false,
                                            is_optional: false,
                                            is_nullable: false }],
                                        body: LambdaBody::Expr(Box::new(Expression::new(
                                            ExprKind::Ident("__e".into()),
                                        ))),
                                        is_async: false,
                                        captures: vec![] });
                                    let len_pred = Expression::new(ExprKind::Lambda {
                                        params: vec![lambda_param("__e")],
                                        body: LambdaBody::Expr(Box::new(Expression::new(
                                            ExprKind::Ternary {
                                                cond: Box::new(Expression::new(ExprKind::Binary {
                                                    op: BinOp::Or,
                                                    left: Box::new(Expression::new(ExprKind::Binary {
                                                        op: BinOp::InstanceOf,
                                                        left: Box::new(Expression::new(ExprKind::Ident(
                                                            "__e".into(),
                                                        ))),
                                                        right: Box::new(Expression::new(ExprKind::Ident(
                                                            "Array".into(),
                                                        ))) })),
                                                    right: Box::new(Expression::new(ExprKind::Binary {
                                                        op: BinOp::StrictEq,
                                                        left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(
                                                            Expression::new(ExprKind::Ident("__e".into())),
                                                        )))),
                                                        right: Box::new(Expression::string("string")) })) })),
                                                then: Box::new(Expression::new(ExprKind::Binary {
                                                    op: BinOp::NotEq,
                                                    left: Box::new(Expression::new(ExprKind::Member {
                                                        object: Box::new(Expression::new(ExprKind::Ident(
                                                            "__e".into(),
                                                        ))),
                                                        field: "length".into(),
                                                        null_safe: false })),
                                                    right: Box::new(Expression::int(0)) })),
                                                else_: Box::new(Expression::new(ExprKind::Ident(
                                                    "__e".into(),
                                                ))) },
                                        ))),
                                        is_async: false,
                                        captures: vec![] });
                                    let inner = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident("filter".into()))),
                                        args: vec![
                                            Argument::positional(ident),
                                            Argument::positional(spread_iterable_expr(desugar_member_reads(
                                                args[1].value.clone(),
                                            ))),
                                        ],
                                        optional: false });
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "filter".into(),
                                        ))),
                                        args: vec![
                                            Argument::positional(len_pred),
                                            Argument::positional(inner),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "filter" if args.len() == 2 && args.iter().all(|a| a.name.is_none()) => {
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "filter".into(),
                                        ))),
                                        args: vec![
                                            Argument::positional(py_callable_expr(args[0].value.clone())),
                                            Argument::positional(spread_iterable_expr(desugar_member_reads(
                                                args[1].value.clone(),
                                            ))),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "iter" if args.len() == 2 && args.iter().all(|a| a.name.is_none()) => {
                                    expr = call_ident(
                                        "__py_iter_sentinel",
                                        vec![
                                            py_callable_expr(args[0].value.clone()),
                                            desugar_member_reads(args[1].value.clone()),
                                        ],
                                    );
                                    continue;
                                }
                                "map" if args.len() >= 2 && args.iter().all(|a| a.name.is_none()) => {
                                    let func = py_callable_expr(args[0].value.clone());
                                    let iterables: Vec<Expression> = args
                                        .iter()
                                        .skip(1)
                                        .map(|a| py_zip_iterable_expr(a.value.clone()))
                                        .collect();
                                    if iterables.len() == 1 {
                                        expr = Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Ident(
                                                "map".into(),
                                            ))),
                                            args: vec![
                                                Argument::positional(func),
                                                Argument::positional(iterables[0].clone()),
                                            ],
                                            optional: false });
                                        continue;
                                    }
                                    let tuple_name = "__py_map_args";
                                    let call_args = (0..iterables.len())
                                        .map(|i| {
                                            Argument::positional(py_index(
                                                Expression::new(ExprKind::Ident(tuple_name.into())),
                                                Expression::int(i as i64),
                                            ))
                                        })
                                        .collect();
                                    let wrapper = py_lambda1(
                                        tuple_name,
                                        Expression::new(ExprKind::Call {
                                        callee: Box::new(func),
                                            args: call_args,
                                            optional: false }),
                                    );
                                    let zipped = Expression::new(ExprKind::Zip {
                                        iterables,
                                        mode: ZipMode::Shortest,
                                        strict: false });
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident("map".into()))),
                                        args: vec![
                                            Argument::positional(wrapper),
                                            Argument::positional(zipped),
                                        ],
                                        optional: false });
                                    continue;
                                }
                                "zip" if args.iter().all(|a| !a.spread) => {
                                    let strict = args.iter().find_map(|a| {
                                        if a.name.as_deref() == Some("strict") {
                                            if matches!(a.value.kind, ExprKind::Lit(Literal::Bool(true))) {
                                                Some(true)
                                            } else {
                                                Some(false)
                                            }
                                        } else {
                                            None
                                        }
                                    });
                                    if strict == Some(true) {
                                        let iterables: Vec<Expression> = args
                                            .iter()
                                            .filter(|a| a.name.as_deref() != Some("strict"))
                                            .map(|a| py_zip_iterable_expr(a.value.clone()))
                                            .collect();
                                        if iterables.len() <= 1 {
                                            expr = Expression::new(ExprKind::Zip {
                                                iterables,
                                                mode: ZipMode::Shortest,
                                                strict: false });
                                            continue;
                                        }
                                        expr = call_ident("__py_zip_strict", iterables);
                                        continue;
                                    } else {
                                        let iterables: Vec<Expression> = args
                                            .iter()
                                            .filter(|a| a.name.as_deref() != Some("strict"))
                                            .map(|a| py_zip_iterable_expr(a.value.clone()))
                                            .collect();
                                        expr = if iterables.is_empty() {
                                            Expression::new(ExprKind::Array(Vec::new()))
                                        } else {
                                            Expression::new(ExprKind::Zip {
                                                iterables,
                                                mode: ZipMode::Shortest,
                                                strict: false })
                                        };
                                        continue;
                                    }
                                }
                                "zip" if args.len() == 1 && args[0].spread && args[0].name.is_none() => {
                                    expr = call_ident(
                                        "__py_zip_spread",
                                        vec![spread_iterable_expr(desugar_member_reads(
                                            args[0].value.clone(),
                                        ))],
                                    );
                                    continue;
                                }
                                "reversed" if args.len() == 1 && args[0].name.is_none() => {
                                    expr = call_ident(
                                        "__py_reversed",
                                        vec![spread_iterable_expr(desugar_member_reads(
                                            args[0].value.clone(),
                                        ))],
                                    );
                                    continue;
                                }
                                "sorted" if args.len() >= 1 => {
                                    // sorted(iterable) → [...iterable].sort()
                                    // sorted(iterable, key=f) → __py_sort_by_key([...iterable], f)
                                    // sorted(..., reverse=True) → … .reverse()
                                    let iterable = desugar_member_reads(args[0].value.clone());
                                    let sorts_tuple_pairs = matches!(
                                        &iterable.kind,
                                        ExprKind::Comprehension {
                                            kind: ComprehensionKind::List,
                                            element,
                                            ..
                                        } if matches!(&element.kind, ExprKind::Tuple(items) if items.len() == 2)
                                    );
                                    let has_reverse =
                                        args.iter().any(|a| a.name.as_deref() == Some("reverse"));
                                    let key_fn = args
                                        .iter()
                                        .find(|a| a.name.as_deref() == Some("key"))
                                        .map(|a| a.value.clone())
                                        .map(wrap_key_ident_in_lambda)
                                        .map(wrap_tuple_key_lambda);
                                    let key_fn = key_fn.or_else(|| {
                                        if sorts_tuple_pairs {
                                            Some(py_lambda1(
                                                "__sk",
                                                py_index(Expression::ident("__sk"), Expression::int(0)),
                                            ))
                                        } else {
                                            None
                                        }
                                    });
                                    // A dict iterates its KEYS, but spreading a
                                    // Map yields [k, v] pairs — route through
                                    // the Python iterate helper first.
                                    let spread_array =
                                        Expression::new(ExprKind::Array(vec![ArrayElement {
                                            key: None,
                                            spread: true,
                                            by_ref: false,
                                            value: call_ident("__py_iter_array__", vec![iterable]) }]));
                                    let sorted = if let Some(key_fn) = key_fn {
                                        call_ident("__py_sort_by_key", vec![spread_array, key_fn])
                                    } else {
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(spread_array),
                                                field: "sort".into(),
                                                null_safe: false })),
                                            args: vec![],
                                            optional: false })
                                    };
                                    expr = if has_reverse {
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(sorted),
                                                field: "reverse".into(),
                                                null_safe: false })),
                                            args: vec![],
                                            optional: false })
                                    } else {
                                        sorted
                                    };
                                    continue;
                                }
                                // `set(iterable)` / `frozenset(iterable)` are left
                                // as plain calls so the profile builtin
                                // (`ecma:set.fromIterable`, which accepts an array,
                                // a string, or nothing) handles every form. A
                                // `New Set` rewrite would need `ecma_new_dispatch`,
                                // which the Python profile does not set, so
                                // `set([1, 2])` failed with "undefined is not
                                // callable".
                                "round" if args.len() == 2 => {
                                    // round(x, n) → Math.round(x * 10**n) / 10**n
                                    let x = args[0].value.clone();
                                    let n = args[1].value.clone();
                                    let factor = Expression::new(ExprKind::Binary {
                                        op: BinOp::Pow,
                                        left: Box::new(Expression::int(10)),
                                        right: Box::new(n) });
                                    let scaled = Expression::new(ExprKind::Binary {
                                        op: BinOp::Mul,
                                        left: Box::new(x),
                                        right: Box::new(factor.clone()) });
                                    let rounded = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "round".into(),
                                        ))),
                                        args: vec![Argument::positional(scaled)],
                                        optional: false });
                                    expr = Expression::new(ExprKind::Binary {
                                        op: BinOp::Div,
                                        left: Box::new(rounded),
                                        right: Box::new(factor) });
                                    continue;
                                }
                                "pow" if args.len() == 3 => {
                                    // pow(base, exp, mod) → pow(base, exp) % mod
                                    let base = args[0].value.clone();
                                    let exp = args[1].value.clone();
                                    let modulus = args[2].value.clone();
                                    let power = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "pow".into(),
                                        ))),
                                        args: vec![
                                            Argument::positional(base),
                                            Argument::positional(exp),
                                        ],
                                        optional: false });
                                    expr = Expression::new(ExprKind::Binary {
                                        op: BinOp::Mod,
                                        left: Box::new(power),
                                        right: Box::new(modulus) });
                                    continue;
                                }
                                _ => {}
                            }
                        }
                        // Python `gen.throw(ExcClass[, ...])` instantiates the class so
                        // the generator's `except` matches an instance (like
                        // `raise ExcClass`). Wrap a bare uppercase-Ident arg.
                        let args = if matches!(&expr.kind, ExprKind::Member { field, .. } if field == "throw")
                            && args.first().is_some_and(|a| {
                                a.name.is_none()
                                    && matches!(&a.value.kind, ExprKind::Ident(n)
                                        if n.chars().next().is_some_and(|c| c.is_ascii_uppercase()))
                            }) {
                            let cls = args[0].value.clone();
                            let ctor_args = args.iter().skip(1).cloned().collect();
                            vec![Argument::positional(Expression::new(ExprKind::Call {
                                callee: Box::new(cls),
                                args: ctor_args,
                                optional: false }))]
                        } else {
                            args
                        };
                        if let Some(rewritten) = rewrite_python_generator_method_call(&expr, &args) {
                            expr = rewritten;
                            continue;
                        }
                        // bytes string-like method with args, e.g.
                        // `b'ab'.replace(b'a', b'x')`, `b'ab'.find(b'b')`.
                        if let ExprKind::Member { object, field, .. } = &expr.kind {
                            if let Some(rewritten) = try_rewrite_bytes_method(object, field, &args)
                            {
                                expr = rewritten;
                                continue;
                            }
                        }
                        // `json.dumps(obj, cls=…, sort_keys=…, indent=…, …)` →
                        // Python-semantics form the json adapter consumes.
                        if let Some(rewritten) = rewrite_json_dumps(&expr, &args) {
                            expr = rewritten;
                            continue;
                        }
                        // `random.NAME(...)` distributions → prelude free fns.
                        if let Some(rewritten) = rewrite_random_call(&expr, &args) {
                            expr = rewritten;
                            continue;
                        }
                        // `dict.fromkeys(keys[, v])` → dict comprehension.
                        if let Some(rewritten) = rewrite_dict_fromkeys(&expr, &args) {
                            expr = rewritten;
                            continue;
                        }
                        // `dict(...)` / `OrderedDict(...)` → dict literal.
                        if let Some(rewritten) = rewrite_dict_construction(&expr, &args) {
                            expr = rewritten;
                            continue;
                        }
                        // `Foo(args)` — construction if `Foo` is a declared class.
                        expr = Expression::new(call_or_new(expr, args));
                    }
                    Rule::identifier => {
                        let field = first_child.as_str().to_string();
                        if field == "__dict__"
                            && !matches!(&expr.kind,
                                ExprKind::Ident(n) if is_imported_module(n))
                        {
                            // Python `obj.__dict__` → the object itself.
                            // Vybe stores instance/class properties in Object.properties,
                            // so ARRAY_GET on the object finds the same keys.
                            // Imported modules keep the Member node — the
                            // desugar pass rebuilds `mod.__dict__` as a real
                            // dict from the namespace object's entries.
                        } else if let Some(global) = prelude_module_class(&expr.kind, &field) {
                            // A prelude module's class (`io.StringIO`,
                            // `configparser.ConfigParser`, …) → the bare global
                            // class, so `mod.Class(...)` CONSTRUCTS directly. As a
                            // method call it would pass the module object as the
                            // constructor's first argument.
                            expr = Expression::new(ExprKind::Ident(global));
                        } else {
                            expr = Expression::new(ExprKind::Member {
                                object: Box::new(expr),
                                field,
                                null_safe: false });
                        }
                    }
                    Rule::subscript => {
                        let index = Expression::new(walk_subscript_expr(
                            children.into_iter().next().unwrap(),
                        )?);
                        let index = python_index_operand(&expr, index);
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(index),
                            null_safe: false });
                    }
                    _ => {
                        // Fallback: try to walk as expression
                        let val = walk_expression(children.into_iter().next().unwrap())?;
                        let val = python_index_operand(&expr, val);
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(val),
                            null_safe: false });
                    }
                }
            }
        }
    }
    Ok(desugar_member_reads(expr).kind)
}

/// Wrap `value` in `[*value]` so a lazy iterable (generator, range, map, …) is
/// drained into an array via the shared `generators.rs` spread machinery before
/// an array/set-based builtin consumes it. Keeps generators cross-language
/// compatible — the drain is the same one JS spread uses.
fn spread_iterable_expr(value: Expression) -> Expression {
    Expression::new(ExprKind::Array(vec![ArrayElement {
        key: None,
        spread: true,
        by_ref: false,
        value }]))
}

fn py_zip_iterable_expr(value: Expression) -> Expression {
    let value = desugar_member_reads(value);
    let is_dict = matches!(py_static_type_name(&value), Some("dict"))
        || matches!(&value.kind, ExprKind::Ident(n) if is_dict_var(n));
    if is_dict {
        call_ident("__py_iter_array__", vec![value])
    } else {
        spread_iterable_expr(value)
    }
}

fn py_known_generator_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => is_generator_var(name),
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => is_generator_func(name),
            ExprKind::FunctionExpr(stmt) => {
                matches!(&stmt.kind, StmtKind::FunctionDecl { is_generator: true, .. })
            }
            _ => false },
        _ => false }
}

fn py_generator_expr_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) if is_generator_func(name) => Some(name.clone()),
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) if is_generator_func(name) => Some(name.clone()),
            ExprKind::FunctionExpr(stmt) => {
                if let StmtKind::FunctionDecl {
                    name,
                    is_generator: true,
                    ..
                } = &stmt.kind
                {
                    Some(name.clone())
                } else {
                    None
                }
            }
            _ => None },
        _ => None }
}

fn rewrite_python_generator_method_call(
    callee_expr: &Expression,
    args: &[Argument],
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee_expr.kind else {
        return None;
    };
    if !py_known_generator_expr(object) {
        return None;
    }
    if args.iter().any(|arg| arg.name.is_some() || arg.spread) {
        return None;
    }
    match field.as_str() {
        "send" if args.len() == 1 => Some(call_ident(
            "__py_gen_send",
            vec![(**object).clone(), args[0].value.clone()],
        )),
        "throw" if args.len() == 1 => Some(call_ident(
            "__py_gen_throw",
            vec![(**object).clone(), args[0].value.clone()],
        )),
        "close" if args.is_empty() => Some(call_ident("__py_gen_close", vec![(**object).clone()])),
        _ => None }
}

/// Normalize Python `print(...)` arguments to the emitter convention
/// `[sep, end, items…]`. The `sep`/`end` keyword args override the defaults
/// (`" "` / `"\n"`); Python's `file`/`flush` keywords are accepted and ignored.
/// Positional items (including `*spread`) keep their original `Argument`.
/// Math functions that return a Python `float` (used by `expr_is_python_float`).
const FLOAT_MATH_FNS: &[&str] = &[
    "sin",
    "cos",
    "tan",
    "asin",
    "acos",
    "atan",
    "atan2",
    "sinh",
    "cosh",
    "tanh",
    "asinh",
    "acosh",
    "atanh",
    "sqrt",
    "pow",
    "exp",
    "log",
    "log2",
    "log10",
    "log1p",
    "expm1",
    "cbrt",
    "degrees",
    "radians",
    "hypot",
    "fabs",
    "fmod",
    "copysign",
    "remainder",
    "dist",
    "fsum",
    "gamma",
    "lgamma",
    "erf",
    "erfc",
    "ldexp",
];

/// The internal builtins the walker lowers Python binary arithmetic to
/// (`+ - * / // % **`). Recognized by the float/bytes-inference passes so their
/// operands are still inspected after lowering.
fn is_py_arith_helper(n: &str) -> bool {
    matches!(
        n,
        "__pyadd__"
            | "__pysub__"
            | "__pymul__"
            | "__pytruediv__"
            | "__pyfloordiv__"
            | "__pymod__"
            | "__pypow__"
    )
}

/// True when an expression is *statically* a Python `float` — a float literal,
/// true division (`/`), `float()`, a float-returning `math.*` call, unary minus
/// of a float, or arithmetic where an operand is a float. Deliberately
/// conservative: never assumes a bare variable or unknown call is a float
/// (that would be the "mark everything" shortcut).
fn expr_is_python_float(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Lit(Literal::Float(_)) => true,
        ExprKind::Binary { op: BinOp::Div, .. } => true,
        ExprKind::Binary {
            op: BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Mod | BinOp::Pow | BinOp::FloorDiv,
            left,
            right } => expr_is_python_float(left) || expr_is_python_float(right),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr } => expr_is_python_float(expr),
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(n) if n == "float" => true,
            ExprKind::Ident(n) if is_float_returning_import(n) => true,
            // `/` lowers to __pytruediv__ and is always float in Python.
            ExprKind::Ident(n) if n == "__pytruediv__" => true,
            // Python arithmetic lowers to __py* helpers — float if an operand is.
            ExprKind::Ident(n) if is_py_arith_helper(n) => {
                args.iter().any(|a| expr_is_python_float(&a.value))
            }
            ExprKind::Member { object, field, .. } => {
                (matches!(&object.kind, ExprKind::Ident(o) if o == "math")
                    && FLOAT_MATH_FNS.contains(&field.as_str()))
                    || (matches!(&object.kind, ExprKind::Ident(o) if o == "statistics")
                        && FLOAT_STATISTICS_FNS.contains(&field.as_str()))
                    || FLOAT_DT_METHODS.contains(&field.as_str())
            }
            _ => false },
        _ => false }
}

/// `datetime` methods CPython documents as returning a float, so they
/// display with a trailing `.0` (`timedelta(hours=1).total_seconds()` is
/// `3600.0`). Named methods, not a blanket rule about the receiver.
const FLOAT_DT_METHODS: &[&str] = &["total_seconds", "timestamp"];

/// `statistics` functions CPython documents as *always* returning a float, so
/// they display with a trailing `.0`. Deliberately not `mean`/`variance`: those
/// return an int for integer data that divides evenly (`mean([42])` is `42`).
const FLOAT_STATISTICS_FNS: &[&str] = &["fmean"];

/// Wrap `value` in `__py_float_repr__(value)` so it displays Python-float-style.
fn wrap_float_repr(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__py_float_repr__".into()))),
        args: vec![Argument::positional(value)],
        optional: false })
}

/// Tier 2 of float display: like `expr_is_python_float` but also treats a bare
/// variable known (from a prior assignment) to hold a float as a float.
fn expr_is_float_ctx(e: &Expression, floats: &HashMap<String, bool>) -> bool {
    match &e.kind {
        ExprKind::Ident(name) => *floats.get(name).unwrap_or(&false),
        ExprKind::Binary { op: BinOp::Div, .. } => true,
        ExprKind::Binary {
            op: BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Mod | BinOp::Pow,
            left,
            right } => expr_is_float_ctx(left, floats) || expr_is_float_ctx(right, floats),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr } => expr_is_float_ctx(expr, floats),
        ExprKind::Call { callee, .. } if matches!(&callee.kind, ExprKind::Ident(n) if n == "__pytruediv__") => {
            true
        }
        ExprKind::Call { callee, args, .. } if matches!(&callee.kind, ExprKind::Ident(n) if is_py_arith_helper(n)) => {
            args.iter().any(|a| expr_is_float_ctx(&a.value, floats))
        }
        _ => expr_is_python_float(e) }
}

/// Wrap bare float-variable arguments of a `print(...)` call so they display
/// Python-float-style. (Direct float expressions were already wrapped during
/// `normalize_python_print_args`; here we catch variables tracked in `floats`.)
fn wrap_float_display_vars(e: &mut Expression, floats: &HashMap<String, bool>) {
    match &mut e.kind {
        ExprKind::Call { callee, args, .. } => {
            if matches!(&callee.kind, ExprKind::Ident(n) if n == "print") {
                // args[0]=sep, args[1]=end, args[2..]=items.
                for a in args.iter_mut().skip(2) {
                    if a.name.is_none()
                        && !a.spread
                        && matches!(&a.value.kind, ExprKind::Ident(name) if *floats.get(name).unwrap_or(&false))
                    {
                        let v = std::mem::replace(&mut a.value, Expression::null());
                        a.value = wrap_float_repr(v);
                    }
                    wrap_float_display_vars(&mut a.value, floats);
                }
            } else {
                wrap_float_display_vars(callee, floats);
                for a in args {
                    wrap_float_display_vars(&mut a.value, floats);
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(value) => {
                        if let ExprKind::Call { callee, args, .. } = &value.kind
                            && matches!(&callee.kind, ExprKind::Ident(n) if n == "str")
                            && args.len() == 1
                            && expr_is_float_ctx(&args[0].value, floats)
                        {
                            let arg = args[0].value.clone();
                            *value = wrap_float_repr(arg);
                        } else if expr_is_float_ctx(value, floats) {
                            let v = std::mem::replace(value, Expression::null());
                            *value = wrap_float_repr(v);
                        } else {
                            wrap_float_display_vars(value, floats);
                        }
                    }
                    InterpolPart::Formatted(value, _) => wrap_float_display_vars(value, floats),
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                wrap_float_display_vars(&mut item.value, floats);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                wrap_float_display_vars(item, floats);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    wrap_float_display_vars(key, floats);
                    wrap_float_display_vars(value, floats);
                }
            }
        }
        ExprKind::Binary { left, right, .. } => {
            wrap_float_display_vars(left, floats);
            wrap_float_display_vars(right, floats);
        }
        ExprKind::Unary { expr, .. } => wrap_float_display_vars(expr, floats),
        ExprKind::Member { object, .. } => wrap_float_display_vars(object, floats),
        ExprKind::Index { object, index, .. } => {
            wrap_float_display_vars(object, floats);
            wrap_float_display_vars(index, floats);
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            wrap_float_display_vars(element, floats);
            for comp_gen in generators {
                wrap_float_display_vars(&mut comp_gen.iter, floats);
                for cond in &mut comp_gen.conditions {
                    wrap_float_display_vars(cond, floats);
                }
            }
        }
        _ => {}
    }
}

/// Post-pass: track which local variables hold floats and wrap float-variable
/// `print` arguments. Function bodies get a fresh scope.
fn apply_float_var_repr(stmts: &mut [Statement], floats: &mut HashMap<String, bool>) {
    for stmt in stmts.iter_mut() {
        match &mut stmt.kind {
            StmtKind::Assign { targets, value , ..} => {
                let is_f = expr_is_float_ctx(value, floats);
                if let [t] = targets.as_slice() {
                    if let ExprKind::Ident(name) = &t.kind {
                        floats.insert(name.clone(), is_f);
                    }
                }
            }
            StmtKind::Expr(e) | StmtKind::Return(Some(e)) => wrap_float_display_vars(e, floats),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                apply_float_var_repr(then_body, floats);
                for (_, b) in elifs.iter_mut() {
                    apply_float_var_repr(b, floats);
                }
                if let Some(b) = else_body {
                    apply_float_var_repr(b, floats);
                }
            }
            StmtKind::While { body, .. } | StmtKind::ForIn { body, .. } => {
                apply_float_var_repr(body, floats)
            }
            StmtKind::For { body, .. } => apply_float_var_repr(body, floats),
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally } => {
                apply_float_var_repr(body, floats);
                for catch in catches {
                    apply_float_var_repr(&mut catch.body, floats);
                }
                if let Some(else_body) = else_body {
                    apply_float_var_repr(else_body, floats);
                }
                if let Some(finally) = finally {
                    apply_float_var_repr(finally, floats);
                }
            }
            StmtKind::Block(b) => apply_float_var_repr(b, floats),
            StmtKind::FunctionDecl { body, .. } => {
                let mut inner = HashMap::new();
                apply_float_var_repr(body, &mut inner);
            }
            _ => {}
        }
    }
}

/// `print(*items, sep=s, end=e, file=f)` with a real `file` target writes
/// `sep.join(map(str, items)) + end` to `f.write(...)` — a stream-backed write
/// (StringIO buffer, `open()` handle, …), never `wasi:logging`. `file=None` or
/// `file=sys.stdout` is the default stdout path, so this returns `None` and the
/// caller emits the normal stream-to-stdout `print`.
fn python_print_file_desugar(args: &[Argument]) -> Option<Expression> {
    let file = args.iter().find(|a| a.name.as_deref() == Some("file"))?;
    // `None` and `sys.stdout` are just the default sink — not a redirect.
    if matches!(file.value.kind, ExprKind::Lit(Literal::Null)) {
        return None;
    }
    if let ExprKind::Member { object, field, .. } = &file.value.kind {
        if field == "stdout" && matches!(&object.kind, ExprKind::Ident(n) if n == "sys") {
            return None;
        }
    }

    let kwarg = |name: &str, default: &str| {
        args.iter()
            .find(|a| a.name.as_deref() == Some(name))
            .map(|a| a.value.clone())
            .unwrap_or_else(|| Expression::string(default))
    };
    let sep = kwarg("sep", " ");
    let end = kwarg("end", "\n");
    let items: Vec<Expression> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| a.value.clone())
        .collect();

    let concat = |a: Expression, b: Expression| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Ident("__pyadd__".into()))),
            args: vec![Argument::positional(a), Argument::positional(b)],
            optional: false })
    };
    // sep.join(str(x) for x in items) + end, built as a left-folded concat.
    let mut acc: Option<Expression> = None;
    for item in items {
        let piece = call_builtin1("str", item);
        acc = Some(match acc {
            None => piece,
            Some(prev) => concat(concat(prev, sep.clone()), piece) });
    }
    let formatted = match acc {
        Some(a) => concat(a, end),
        None => end };

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(file.value.clone()),
            field: "write".into(),
            null_safe: false })),
        args: vec![Argument::positional(formatted)],
        optional: false }))
}

fn normalize_python_print_args(raw: Vec<Argument>) -> Vec<Argument> {
    let mut sep = Argument::positional(Expression::string(" "));
    let mut end = Argument::positional(Expression::string("\n"));
    let mut items = Vec::new();
    for a in raw {
        match a.name.as_deref() {
            Some("sep") => sep = Argument::positional(a.value),
            Some("end") => end = Argument::positional(a.value),
            Some("file") | Some("flush") => {}
            _ => {
                // Format statically-known floats Python-style (`4.0`, not `4`).
                // Bytes display is handled at runtime in `emit_py_repr` via
                // `arraybuffer.isView`, so no static wrapping is needed here.
                let value = desugar_member_reads(a.value);
                if is_counter_expr(&value) {
                    items.push(Argument::positional(call_ident("__py_counter_repr", vec![value])));
                } else if expr_is_python_float(&value) {
                    items.push(Argument::positional(wrap_float_repr(value)));
                } else {
                    items.push(Argument {
                        value,
                        ..a
                    });
                }
            }
        }
    }
    let mut out = Vec::with_capacity(items.len() + 2);
    out.push(sep);
    out.push(end);
    out.extend(items);
    out
}

/// `print(*items, sep=s, end=e)` with a spread positional argument.
///
/// `emit_print` takes a statically-fixed argument count, so it cannot expand a
/// spread (`*seq`) whose length is only known at runtime — and by the time the
/// items reach the stack, `print([1,2,3])` and `print(*[1,2,3])` are
/// indistinguishable. The intent only exists here, so we desugar to the
/// runtime-join form Python's `print` is defined by:
///
/// ```text
/// print(*a, x, sep=s, end=e)  →  print(s.join([str(v) for v in [*a, x]]), end=e)
/// ```
///
/// Returns a replacement argument list — a single joined positional item plus
/// the preserved `end` keyword — that the ordinary `normalize_python_print_args`
/// fixed-argc path then handles. Returns `None` when no positional argument is a
/// spread (the common case), leaving every non-spread `print` byte-identical.
fn python_print_spread_desugar(args: &[Argument]) -> Option<Vec<Argument>> {
    let has_positional_spread = args.iter().any(|a| a.name.is_none() && a.spread);
    if !has_positional_spread {
        return None;
    }

    let kwarg = |name: &str| args.iter().find(|a| a.name.as_deref() == Some(name));
    let sep = kwarg("sep")
        .map(|a| a.value.clone())
        .unwrap_or_else(|| Expression::string(" "));

    // Flatten every positional argument into one array, preserving spreads so
    // `[*a, x]` expands instead of nesting.
    let elements: Vec<ArrayElement> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| ArrayElement {
            key: None,
            value: a.value.clone(),
            spread: a.spread,
            by_ref: false })
        .collect();
    let flattened = Expression::new(ExprKind::Array(elements));

    // `[str(v) for v in <flattened>]`
    let stringified = Expression::new(ExprKind::Comprehension {
        kind: ComprehensionKind::List,
        element: Box::new(call_builtin1("str", Expression::new(ExprKind::Ident(
            "__print_item".into(),
        )))),
        generators: vec![ComprehensionGen {
            target: Expression::new(ExprKind::Ident("__print_item".into())),
            iter: flattened,
            conditions: Vec::new(),
            is_async: false }] });

    // `<list>.join(sep)` — the swapped `array.join(delim)` convention the
    // compiler expects (the source-level `delim.join(array)` swap does not run
    // on synthesized nodes).
    let joined = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(stringified),
            field: "join".into(),
            null_safe: false })),
        args: vec![Argument::positional(sep)],
        optional: false });

    // Replacement args: one joined positional (no spread) plus the preserved
    // `end` keyword. `normalize_python_print_args` reshapes this to the
    // fixed-argc [sep, end, item] convention. `sep` is already applied by the
    // join, and a single item never re-applies it.
    let mut print_args = vec![Argument::positional(joined)];
    if let Some(end) = kwarg("end") {
        print_args.push(Argument {
            value: end.value.clone(),
            name: Some("end".into()),
            by_ref: false,
            spread: false });
    }
    Some(print_args)
}

fn walk_call_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut args = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::call_arg {
            // The `*` / `**` prefixes are silent literals in the grammar, so
            // check the call_arg's own text rather than a child token.
            let arg_text = p.as_str().trim_start();
            let is_starstar = arg_text.starts_with("**");
            let is_star = !is_starstar && arg_text.starts_with('*');
            let mut ci: Vec<Pair<Rule>> = p.into_inner().collect();
            if ci.is_empty() {
                continue;
            }

            if is_starstar {
                // **kwargs — spread of a mapping into keyword arguments.
                let val = walk_expression(ci.pop().unwrap())?;
                let val = match val.kind {
                    ExprKind::Spread(inner) => *inner,
                    _ => val };
                args.push(Argument {
                    value: val,
                    name: None,
                    by_ref: false,
                    spread: true });
            } else if is_star {
                // *args — positional spread expansion.
                let val = walk_expression(ci.pop().unwrap())?;
                let val = match val.kind {
                    ExprKind::Spread(inner) => *inner,
                    _ => val };
                args.push(Argument {
                    value: val,
                    name: None,
                    by_ref: false,
                    spread: true });
            } else if ci.len() >= 2 && ci[0].as_rule() == Rule::identifier {
                // Check if it's keyword=value: identifier followed by expression
                // If there's an "=" between them
                let name = ci[0].as_str().to_string();
                let val = walk_expression(ci.pop().unwrap())?;
                args.push(Argument {
                    value: val,
                    name: Some(name),
                    by_ref: false,
                    spread: false });
            } else if ci[0].as_rule() == Rule::comp_for_arg {
                // Generator expression as argument
                let val = walk_expression(ci.remove(0))?;
                args.push(Argument::positional(val));
            } else {
                // `*args` parses as a `star_expr` → `ExprKind::Spread`; unwrap it
                // and flag the argument as spread so the call expands it.
                let val = walk_expression(ci.remove(0))?;
                if let ExprKind::Spread(inner) = val.kind {
                    args.push(Argument {
                        value: *inner,
                        name: None,
                        by_ref: false,
                        spread: true });
                } else {
                    args.push(Argument::positional(val));
                }
            }
        }
    }
    Ok(args)
}

// ── Subscript ───────────────────────────────────────────────────────────────

/// Wrap a scalar subscript index in the from-end offset normalizer
/// `__py_from_end__` so `a[-1]` reads one-from-the-end (like C#'s `arr[^N]`).
/// Skips keys that can never be a from-end offset: string literals (dict keys),
/// non-negative integer literals, and slices/ranges (the slice path already
/// offsets from the end). The normalizer is a runtime no-op unless the index is
/// a negative number on a sequence, so dict lookups stay direct.
fn python_index_operand(object: &Expression, index: Expression) -> Expression {
    match &index.kind {
        ExprKind::Lit(Literal::Str(_)) => return index,
        ExprKind::Lit(Literal::Int(n)) if *n >= 0 => return index,
        ExprKind::Slice { .. } | ExprKind::Range { .. } => return index,
        _ => {}
    }
    // `__py_from_end__` re-reads the base to compute its length, so the base is
    // cloned into the call. At this point the postfix chain hasn't been desugared
    // yet (that happens once at the end of `walk_postfix`), so a property read
    // like `obj.parts` is still a raw `Member`. The outer subscript base is
    // desugared later, but this embedded clone would be missed — desugar it here
    // so a property/attribute base (`obj.parts[-1]`) reads through the same
    // subscript form as everywhere else instead of a raw property access.
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__py_from_end__")),
        args: vec![
            Argument::positional(desugar_member_reads(object.clone())),
            Argument::positional(index),
        ],
        optional: false })
}

/// Is this subscript index a slice rather than a scalar key?
fn index_is_slice(index: &Expression) -> bool {
    matches!(
        index.kind,
        ExprKind::Slice { .. } | ExprKind::Range { .. }
    )
}

/// Desugar the bound expressions inside a slice/range index. `desugar_member_reads`
/// falls through to its catch-all for `Slice`/`Range`, so a bound like
/// `l[obj.start:2]` would otherwise keep a raw `Member` that never gets rewritten
/// into the subscript form the rest of the walker produces.
fn desugar_slice_bounds(index: Expression) -> Expression {
    let desugar_opt =
        |b: Option<Box<Expression>>| b.map(|e| Box::new(desugar_member_reads(*e)));
    match index.kind {
        ExprKind::Slice { lower, upper, step } => Expression::new(ExprKind::Slice {
            lower: desugar_opt(lower),
            upper: desugar_opt(upper),
            step: desugar_opt(step) }),
        ExprKind::Range {
            start,
            end,
            inclusive } => Expression::new(ExprKind::Range {
            start: Box::new(desugar_member_reads(*start)),
            end: Box::new(desugar_member_reads(*end)),
            inclusive }),
        _ => index }
}

fn walk_subscript_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let text = pair.as_str().trim();
    if pair.as_rule() == Rule::subscript_item && text.contains(':') {
        let mut exprs = pair
            .into_inner()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?
            .into_iter();
        let mut parts = text.split(':').map(str::trim);
        let lower = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice lower bound")?)) };
        let upper = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice upper bound")?)) };
        let step = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice step")?)) };
        return Ok(ExprKind::Slice { lower, upper, step });
    }

    let items: Vec<Pair<Rule>> = pair.into_inner().collect();
    if items.len() == 1 {
        return walk_expr_kind(items.into_iter().next().unwrap());
    }
    // Multiple subscript items → tuple
    let exprs = items
        .into_iter()
        .map(walk_expression)
        .collect::<Result<Vec<_>, _>>()?;
    Ok(ExprKind::Tuple(exprs))
}

// ── Primary ─────────────────────────────────────────────────────────────────

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }
    // Multiple children in primary — could be parenthesized expr, list, dict, etc.
    if inner.is_empty() {
        return Ok(ExprKind::Tuple(Vec::new())); // empty tuple ()
    }
    // Check what we have
    let first = &inner[0];
    match first.as_rule() {
        Rule::expression_list => {
            // Parenthesized expression or tuple
            let expr = walk_expr_list(inner.remove(0))?;
            Ok(expr.kind)
        }
        Rule::list_inner => walk_list_inner(inner.remove(0)),
        Rule::dict_or_set_inner => walk_dict_or_set(inner.remove(0)),
        _ => walk_expr_kind(inner.remove(0)) }
}

fn walk_list_inner(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.is_empty() {
        return Ok(ExprKind::Array(Vec::new()));
    }

    // Check for comprehension
    let has_comp = inner.iter().any(|p| p.as_rule() == Rule::comp_clause);
    if has_comp {
        let element = walk_expression(inner.remove(0))?;
        let generators = inner
            .into_iter()
            .filter(|p| p.as_rule() == Rule::comp_clause)
            .map(walk_comp_clause)
            .collect::<Result<Vec<_>, _>>()?;
        return Ok(ExprKind::Comprehension {
            kind: ComprehensionKind::List,
            element: Box::new(element),
            generators });
    }

    // Normal list. `*x` elements walk to `ExprKind::Spread(x)`; unwrap them and
    // set the `spread` flag so `[*a, *b]` flattens instead of nesting.
    let elements = inner
        .into_iter()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(|p| -> Result<ArrayElement, String> {
            let val = walk_expression(p)?;
            if let ExprKind::Spread(inner) = val.kind {
                Ok(ArrayElement {
                    key: None,
                    value: *inner,
                    spread: true,
                    by_ref: false })
            } else {
                Ok(ArrayElement {
                    key: None,
                    value: val,
                    spread: false,
                    by_ref: false })
            }
        })
        .collect::<Result<Vec<_>, String>>()?;
    Ok(ExprKind::Array(elements))
}

fn walk_dict_or_set(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let text = pair.as_str().trim();
    if text.is_empty() {
        return Ok(ExprKind::Object(Vec::new())); // empty dict {}
    }

    let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.is_empty() {
        return Ok(ExprKind::Object(Vec::new()));
    }

    // ── Set comprehension: expression ~ set_comp_or_rest(comp_clause+) ──
    // e.g. {x % 2 for x in range(6)}
    if inner.len() >= 2 && inner[1].as_rule() == Rule::set_comp_or_rest {
        let set_inner: Vec<Pair<Rule>> = inner[1].clone().into_inner().collect();
        let has_comp = set_inner.iter().any(|p| p.as_rule() == Rule::comp_clause);
        if has_comp {
            let element = walk_expression(inner[0].clone())?;
            let generators = set_inner
                .into_iter()
                .filter(|p| p.as_rule() == Rule::comp_clause)
                .map(walk_comp_clause)
                .collect::<Result<Vec<_>, _>>()?;
            return Ok(ExprKind::Comprehension {
                kind: ComprehensionKind::Set,
                element: Box::new(element),
                generators });
        }
    }

    // ── Dict comprehension: expr ~ expr ~ dict_comp_or_rest(comp_clause+) ──
    // e.g. {x: x * x for x in range(4)}
    if inner.len() >= 3 && inner[2].as_rule() == Rule::dict_comp_or_rest {
        let comp_inner: Vec<Pair<Rule>> = inner[2].clone().into_inner().collect();
        let has_comp = comp_inner.iter().any(|p| p.as_rule() == Rule::comp_clause);
        if has_comp
            && is_expression_rule(inner[0].as_rule())
            && is_expression_rule(inner[1].as_rule())
        {
            let key = walk_expression(inner[0].clone())?;
            let val = walk_expression(inner[1].clone())?;
            // Encode key-value as a 2-element array so the compiler can unpack it.
            let element = Expression::new(ExprKind::Array(vec![
                ArrayElement {
                    key: None,
                    spread: false,
                    by_ref: false,
                    value: key },
                ArrayElement {
                    key: None,
                    spread: false,
                    by_ref: false,
                    value: val },
            ]));
            let generators = comp_inner
                .into_iter()
                .filter(|p| p.as_rule() == Rule::comp_clause)
                .map(walk_comp_clause)
                .collect::<Result<Vec<_>, _>>()?;
            return Ok(ExprKind::Comprehension {
                kind: ComprehensionKind::Dict,
                element: Box::new(element),
                generators });
        }
    }

    // ── Dict literal or set literal ──────────────────────────────────────
    // Python quirk: empty `{}` is an empty DICT, not a set (`set()` is the
    // empty set). Without this, `{}` became `ExprKind::Set([])` and `d[k]=v`
    // never created enumerable object properties.
    let mut is_dict = inner.is_empty();
    for p in &inner {
        match p.as_rule() {
            Rule::dict_comp_or_rest | Rule::dict_rest | Rule::dict_entry => is_dict = true,
            _ => {}
        }
    }

    if is_dict {
        let mut props = Vec::new();
        let mut i = 0;
        while i < inner.len() {
            match inner[i].as_rule() {
                Rule::dict_comp_or_rest | Rule::dict_rest => {
                    for de in inner[i].clone().into_inner() {
                        if de.as_rule() == Rule::dict_entry {
                            let is_spread = de.as_str().trim_start().starts_with("**");
                            let entry_inner: Vec<Pair<Rule>> = de.into_inner().collect();
                            if is_spread {
                                if let Some(expr) = entry_inner.first() {
                                    props.push(ObjectProperty::Spread(walk_expression(
                                        expr.clone(),
                                    )?));
                                }
                            } else if entry_inner.len() >= 2 {
                                let key = walk_expression(entry_inner[0].clone())?;
                                let val = walk_expression(entry_inner[1].clone())?;
                                props.push(ObjectProperty::KeyValue { key, value: val });
                            }
                        }
                    }
                }
                _ if is_expression_rule(inner[i].as_rule()) => {
                    let key = walk_expression(inner[i].clone())?;
                    if i == 0 && text.starts_with("**") {
                        props.push(ObjectProperty::Spread(key));
                        i += 1;
                        continue;
                    }
                    i += 1;
                    if i < inner.len() && is_expression_rule(inner[i].as_rule()) {
                        let val = walk_expression(inner[i].clone())?;
                        props.push(ObjectProperty::KeyValue { key, value: val });
                    }
                }
                _ => {}
            }
            i += 1;
        }
        return Ok(ExprKind::Object(props));
    }

    // Set literal: {1, 2, 3}
    let mut elements = Vec::new();
    for item in inner {
        match item.as_rule() {
            rule if is_expression_rule(rule) => elements.push(walk_expression(item)?),
            Rule::set_comp_or_rest => {
                for part in item.into_inner() {
                    if is_expression_rule(part.as_rule()) {
                        elements.push(walk_expression(part)?);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(ExprKind::Set(elements))
}

/// Destructure a tuple loop target `(a, b, …)` bound to a temp into
/// `a = tmp[0]; b = tmp[1]; …` statements, prepended to a genexpr loop body.
fn destructure_comp_target(target: &Expression, tmp: &str) -> Vec<Statement> {
    let mut out = Vec::new();
    if let ExprKind::Tuple(parts) = &target.kind {
        for (i, part) in parts.iter().enumerate() {
            if let ExprKind::Ident(name) = &part.kind {
                out.push(Statement::new(StmtKind::Assign {
                    targets: vec![Expression::new(ExprKind::Ident(name.clone()))],
                    value: Expression::new(ExprKind::Index {
                        object: Box::new(Expression::ident(tmp)),
                        index: Box::new(Expression::int(i as i64)),
                        null_safe: false }), by_ref: false }));
            }
        }
    }
    out
}

/// Lower a Python generator expression `(element for t in iter if cond …)` into
/// an immediately-invoked generator function `(function* () { … yield element …
/// })()` so it stays lazy — driven by `next()` through the shared generator
/// machinery — instead of the eager comprehension path that materializes the
/// entire iterator.
fn lower_generator_expression(element: Expression, generators: Vec<ComprehensionGen>) -> ExprKind {
    // Innermost body: `yield element`.
    let mut stmts: Vec<Statement> = vec![Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Yield(Some(Box::new(element))),
    )))];

    // Nest the clauses from the innermost generator outward.
    for clause in generators.into_iter().rev() {
        // Apply this clause's `if` filters (innermost first).
        for cond in clause.conditions.into_iter().rev() {
            stmts = vec![Statement::new(StmtKind::If {
                cond,
                then_body: stmts,
                elifs: Vec::new(),
                else_body: None })];
        }
        let (var, body) = match &clause.target.kind {
            ExprKind::Ident(name) => (name.clone(), stmts),
            _ => {
                let tmp = "__ge_elem".to_string();
                let mut body = destructure_comp_target(&clause.target, &tmp);
                body.extend(stmts);
                (tmp, body)
            }
        };
        stmts = vec![Statement::new(StmtKind::ForIn {
            var,
            key: None,
            iter: clause.iter,
            body,
            of: true,
            else_body: None,
            is_async: clause.is_async })];
    }

    let gen_fn = Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: Vec::new(),
            return_type: None,
            body: stmts,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: true,
            is_sub: false },
    ))));
    ExprKind::Call {
        callee: Box::new(gen_fn),
        args: Vec::new(),
        optional: false }
}

/// Lower `range(stop)` / `range(start, stop)` / `range(start, stop, step)` into
/// a lazy generator IIFE so it never materializes the whole sequence — the same
/// `generators.rs` stack-switching engine the generator expression uses. Only
/// the bare positional forms are lowered; anything else keeps the profile
/// builtin.
fn walk_comp_clause(pair: Pair<Rule>) -> Result<ComprehensionGen, String> {
    let mut target = Expression::new(ExprKind::Ident("_".into()));
    let mut iter = Expression::new(ExprKind::Lit(Literal::Null));
    let mut conditions = Vec::new();
    let mut is_async = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::async_kw => is_async = true,
            Rule::target_list => {
                // `for a, b in …` unpacks each element; represent multi-name
                // targets as a tuple so the comprehension compiler destructures
                // them (a bare `Ident("a, b")` would bind the whole pair to one
                // variable named "a, b").
                let raw = p.as_str().trim().to_string();
                let names: Vec<String> = raw
                    .split(',')
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty())
                    .collect();
                target = if names.len() > 1 {
                    Expression::new(ExprKind::Tuple(
                        names
                            .into_iter()
                            .map(|n| Expression::new(ExprKind::Ident(n)))
                            .collect(),
                    ))
                } else {
                    Expression::new(ExprKind::Ident(raw))
                };
            }
            Rule::in_kw => {}
            Rule::comp_if => {
                for ci in p.into_inner() {
                    if is_expression_rule(ci.as_rule()) {
                        conditions.push(walk_expression(ci)?);
                    }
                }
            }
            _ if is_expression_rule(p.as_rule()) => {
                iter = walk_expression(p)?;
            }
            _ => {}
        }
    }

    // Wrap string literals in [...s] so comprehensions iterate chars
    if matches!(iter.kind, ExprKind::Lit(Literal::Str(_))) {
        iter = Expression::new(ExprKind::Array(vec![ArrayElement {
            key: None,
            spread: true,
            by_ref: false,
            value: iter }]));
    }

    Ok(ComprehensionGen {
        target,
        iter,
        conditions,
        is_async })
}

// ── Lambda ──────────────────────────────────────────────────────────────────

fn walk_lambda(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut params = Vec::new();
    let mut body_expr = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::lambda_params => {
                for lp in p.into_inner() {
                    if lp.as_rule() == Rule::lambda_param {
                        let mut name = String::new();
                        let mut default = None;
                        let mut is_rest = false;
                        let mut is_kwargs = false;
                        for c in lp.into_inner() {
                            match c.as_rule() {
                                Rule::identifier => name = c.as_str().to_string(),
                                _ if c.as_str() == "**" => is_kwargs = true,
                                _ if c.as_str() == "*" => is_rest = true,
                                _ => default = Some(walk_expression(c)?) }
                        }
                        if !name.is_empty() {
                            params.push(Param {
                                name,
                                type_hint: None,
                                is_optional: default.is_some(),
                                default,
                                pass_by: PassBy::Value,
                                is_rest,
                                is_kwargs,
                                is_nullable: false });
                        }
                    }
                }
            }
            _ if is_expression_rule(p.as_rule()) => {
                body_expr = Some(walk_expression(p)?);
            }
            _ => {}
        }
    }

    Ok(ExprKind::Lambda {
        params,
        body: LambdaBody::Expr(Box::new(body_expr.unwrap_or(Expression::null()))),
        is_async: false,
        captures: Vec::new() })
}

// ── Yield ───────────────────────────────────────────────────────────────────

fn walk_yield(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut is_from = false;
    let mut expr = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::yield_kw => {}
            Rule::yield_from_kw => is_from = true,
            _ if is_expression_rule(p.as_rule()) => expr = Some(walk_expression(p)?),
            Rule::expression_list => expr = Some(walk_expr_list(p)?),
            _ => {}
        }
    }

    if is_from {
        Ok(ExprKind::YieldFrom(Box::new(
            expr.unwrap_or(Expression::null()),
        )))
    } else {
        Ok(ExprKind::Yield(expr.map(Box::new)))
    }
}

// ── F-string ────────────────────────────────────────────────────────────────

/// Expand a compile-time `"template".format(args)` call into an interpolation
/// AST, reusing `str`/`repr` and the `%`-formatting path for `{:spec}` fields.
/// Returns `None` when the template or a field uses something we can't expand
/// statically (`**kwargs`/`*args` spreads, nested `{}` in a spec, Python-only
/// specs like `,`/`^`/`%`), so the caller falls through to the current behavior
/// — no regression. Only valid for a string-literal receiver.
fn expand_str_format(template: &str, args: &[Argument]) -> Option<Expression> {
    // Spreads (`*args`, `**kwargs`) can't be resolved to fields statically.
    if args.iter().any(|a| a.spread) {
        return None;
    }
    let positionals: Vec<Expression> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| a.value.clone())
        .collect();

    let chars: Vec<char> = template.chars().collect();
    let mut parts: Vec<InterpolPart> = Vec::new();
    let mut text = String::new();
    let mut auto_idx = 0usize;
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '{' {
            if i + 1 < chars.len() && chars[i + 1] == '{' {
                text.push('{');
                i += 2;
                continue;
            }
            if !text.is_empty() {
                parts.push(InterpolPart::Text(std::mem::take(&mut text)));
            }
            // Scan to the matching '}'. A nested '{' inside the field (dynamic
            // spec like `{:{w}}`) is unsupported → bail.
            let mut j = i + 1;
            let mut body = String::new();
            while j < chars.len() && chars[j] != '}' {
                if chars[j] == '{' {
                    return None;
                }
                body.push(chars[j]);
                j += 1;
            }
            if j >= chars.len() {
                return None; // unterminated field
            }
            i = j + 1;
            let field_expr = expand_format_field(&body, &positionals, args, &mut auto_idx)?;
            parts.push(InterpolPart::Expr(field_expr));
        } else if c == '}' {
            if i + 1 < chars.len() && chars[i + 1] == '}' {
                text.push('}');
                i += 2;
                continue;
            }
            return None; // lone '}'
        } else {
            text.push(c);
            i += 1;
        }
    }
    if !text.is_empty() {
        parts.push(InterpolPart::Text(text));
    }
    Some(Expression::new(ExprKind::Interpolation(parts)))
}

/// Build one replacement field: `[name][!conv][:spec]`.
fn expand_format_field(
    body: &str,
    positionals: &[Expression],
    args: &[Argument],
    auto_idx: &mut usize,
) -> Option<Expression> {
    let (head, spec) = match body.find(':') {
        Some(p) => (&body[..p], Some(&body[p + 1..])),
        None => (body, None) };
    let (name_part, conv) = match head.find('!') {
        Some(p) => (&head[..p], Some(&head[p + 1..])),
        None => (head, None) };

    let base = resolve_format_value(name_part, positionals, args, auto_idx)?;

    // Apply the `!r` / `!s` conversion first (Python order: convert, then spec).
    let converted = match conv {
        None => base,
        Some("r") | Some("a") => call_builtin1("repr", base),
        Some("s") => call_builtin1("str", base),
        Some(_) => return None };

    if let Some(spec) = spec {
        if !spec.is_empty() {
            // Native Python formatting first (handles what printf can't, e.g.
            // `^` centre); otherwise fall back to the `fmt % value` printf path.
            if let Some(native) = expand_python_format_spec(converted.clone(), spec) {
                return Some(native);
            }
            let fmt = format_spec_to_printf(spec)?;
            return Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__pymod__".into()))),
                args: vec![
                    Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(fmt)))),
                    Argument::positional(converted),
                ],
                optional: false }));
        }
    }

    // No spec: str() unless a conversion already produced a string.
    Some(if conv.is_none() {
        call_builtin1("str", converted)
    } else {
        converted
    })
}

fn call_builtin1(name: &str, arg: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
        args: vec![Argument::positional(arg)],
        optional: false })
}

/// Resolve a field name (`""` auto, `"0"` positional, `"name"` keyword) plus
/// `[k]` / `.attr` accessors into a value expression.
fn resolve_format_value(
    name_part: &str,
    positionals: &[Expression],
    args: &[Argument],
    auto_idx: &mut usize,
) -> Option<Expression> {
    let acc_start = name_part.find(['[', '.']).unwrap_or(name_part.len());
    let base_name = &name_part[..acc_start];
    let mut rest = &name_part[acc_start..];

    let mut expr = if base_name.is_empty() {
        let idx = *auto_idx;
        *auto_idx += 1;
        positionals.get(idx)?.clone()
    } else if let Ok(idx) = base_name.parse::<usize>() {
        positionals.get(idx)?.clone()
    } else {
        args.iter()
            .find(|a| a.name.as_deref() == Some(base_name))
            .map(|a| a.value.clone())?
    };

    while !rest.is_empty() {
        if let Some(stripped) = rest.strip_prefix('[') {
            let end = stripped.find(']')?;
            let key = &stripped[..end];
            let index_expr = if let Ok(n) = key.parse::<i64>() {
                Expression::new(ExprKind::Lit(Literal::Int(n)))
            } else {
                Expression::new(ExprKind::Lit(Literal::Str(key.to_string())))
            };
            expr = Expression::new(ExprKind::Index {
                object: Box::new(expr),
                index: Box::new(index_expr),
                null_safe: false });
            rest = &stripped[end + 1..];
        } else if let Some(stripped) = rest.strip_prefix('.') {
            let end = stripped.find(['[', '.']).unwrap_or(stripped.len());
            let attr = &stripped[..end];
            if attr.is_empty() {
                return None;
            }
            expr = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: attr.to_string(),
                null_safe: false });
            rest = &stripped[end..];
        } else {
            return None;
        }
    }
    Some(expr)
}

/// Translate a Python format spec into an equivalent printf conversion
/// (`>3` → `%3s`, `.2f` → `%.2f`, `+d` → `%+d`). Returns `None` for
/// Python-only specs printf can't express (fill char, `^` center, `,`/`_`
/// grouping, `%`/`n` types) so the caller bails cleanly.
fn format_spec_to_printf(spec: &str) -> Option<String> {
    let chars: Vec<char> = spec.chars().collect();
    let mut idx = 0;

    // [[fill]align]
    let mut align: Option<char> = None;
    let mut fill: Option<char> = None;
    if chars.len() >= 2 && matches!(chars[1], '<' | '>' | '^' | '=') {
        fill = Some(chars[0]);
        align = Some(chars[1]);
        idx = 2;
    } else if !chars.is_empty() && matches!(chars[0], '<' | '>' | '^' | '=') {
        align = Some(chars[0]);
        idx = 1;
    }
    // printf can't do a custom fill char, center, or pad-after-sign.
    if let Some(f) = fill {
        if f != ' ' && f != '0' {
            return None;
        }
    }
    if matches!(align, Some('^') | Some('=')) {
        return None;
    }

    let mut flags = String::new();
    if align == Some('<') {
        flags.push('-');
    }
    // sign
    if idx < chars.len() && matches!(chars[idx], '+' | '-' | ' ') {
        flags.push(chars[idx]);
        idx += 1;
    }
    // alternate form
    if idx < chars.len() && chars[idx] == '#' {
        flags.push('#');
        idx += 1;
    }
    // zero-pad
    if idx < chars.len() && chars[idx] == '0' {
        if !flags.contains('0') {
            flags.push('0');
        }
        idx += 1;
    }
    // width
    let mut width = String::new();
    while idx < chars.len() && chars[idx].is_ascii_digit() {
        width.push(chars[idx]);
        idx += 1;
    }
    // grouping (Python-only)
    if idx < chars.len() && matches!(chars[idx], ',' | '_') {
        return None;
    }
    // precision
    let mut precision = String::new();
    if idx < chars.len() && chars[idx] == '.' {
        precision.push('.');
        idx += 1;
        while idx < chars.len() && chars[idx].is_ascii_digit() {
            precision.push(chars[idx]);
            idx += 1;
        }
    }
    // type
    let ty = if idx < chars.len() {
        let t = chars[idx];
        idx += 1;
        t
    } else {
        's'
    };
    if idx != chars.len() {
        return None; // trailing junk
    }
    let conv = match ty {
        'd' | 's' | 'x' | 'X' | 'o' | 'b' | 'e' | 'E' | 'f' | 'F' | 'g' | 'G' | 'c' => ty,
        _ => return None, // 'n', '%', etc.
    };

    let mut out = String::from("%");
    if fill == Some('0') && !flags.contains('0') {
        out.push('0');
    }
    out.push_str(&flags);
    out.push_str(&width);
    out.push_str(&precision);
    out.push(conv);
    Some(out)
}

/// Native Python format for spec shapes C-printf can't express — Python's format
/// mini-language is its own sprintf dialect, not C printf, so it composes Python
/// string methods and dedicated emitter primitives (`__py_fmt_*`, built on the
/// ECMA number ops) directly rather than translating to `%…`.
///
/// Owns: alignment `[[fill]align]width` (incl. `^` centre, which printf can't
/// express), scientific `e`/`E`, percent `%`, and thousands grouping `,`/`_`.
/// Returns `None` for any shape it doesn't yet own (plain `d`/`f`/`x`/`o`/`b`/`g`
/// without a grouping/percent/scientific twist), so the caller falls back to
/// `format_spec_to_printf`.
fn expand_python_format_spec(value: Expression, spec: &str) -> Option<Expression> {
    let chars: Vec<char> = spec.chars().collect();
    let mut i = 0;

    // [[fill]align]
    let (mut fill, mut align) = (' ', None);
    if chars.len() >= 2 && matches!(chars[1], '<' | '>' | '^' | '=') {
        fill = chars[0];
        align = Some(chars[1]);
        i = 2;
    } else if chars
        .first()
        .is_some_and(|c| matches!(c, '<' | '>' | '^' | '='))
    {
        align = Some(chars[0]);
        i = 1;
    }
    if align == Some('=') {
        return None; // sign-aware padding: not owned natively yet
    }
    // sign
    let sign = if chars.get(i).is_some_and(|c| matches!(c, '+' | '-' | ' ')) {
        let s = chars[i];
        i += 1;
        Some(s)
    } else {
        None
    };
    if chars.get(i) == Some(&'#') {
        return None; // alternate form → printf
    }
    let zero = chars.get(i) == Some(&'0');
    if zero {
        i += 1;
    }
    // width
    let mut width_s = String::new();
    while chars.get(i).is_some_and(|c| c.is_ascii_digit()) {
        width_s.push(chars[i]);
        i += 1;
    }
    let width: Option<i64> = if width_s.is_empty() {
        None
    } else {
        width_s.parse().ok()
    };
    // grouping
    let grouping = if chars.get(i).is_some_and(|c| matches!(c, ',' | '_')) {
        let g = chars[i];
        i += 1;
        Some(g)
    } else {
        None
    };
    // precision
    let mut precision: Option<i64> = None;
    if chars.get(i) == Some(&'.') {
        i += 1;
        let mut p = String::new();
        while chars.get(i).is_some_and(|c| c.is_ascii_digit()) {
            p.push(chars[i]);
            i += 1;
        }
        precision = p.parse().ok();
    }
    // type
    let ty = chars.get(i).copied();
    if ty.is_some() {
        i += 1;
    }
    if i != chars.len() {
        return None; // trailing junk
    }

    let int_lit = |n: i64| Expression::new(ExprKind::Lit(Literal::Int(n)));
    let str_lit = |s: &str| Expression::new(ExprKind::Lit(Literal::Str(s.to_string())));
    let call2 = |name: &str, a: Expression, b: Expression| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
            args: vec![Argument::positional(a), Argument::positional(b)],
            optional: false })
    };

    // A native numeric base (scientific/percent/grouping) right-aligns by
    // default; a plain string base's default alignment is runtime-type-dependent
    // (ints right, strings left) so we leave bare-width string cases to printf.
    let numeric_base = matches!(ty, Some('e') | Some('E') | Some('%')) || grouping.is_some();

    // Base string per type. Grouping only supported on integer-ish bases here
    // (no fractional grouping yet).
    let base: Expression = match ty {
        Some('e') | Some('E') => {
            let sci = call2("__py_fmt_sci", value, int_lit(precision.unwrap_or(6)));
            if ty == Some('E') {
                // Uppercase only flips the exponent `e`; digits/sign are unaffected.
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(sci),
                        field: "upper".into(),
                        null_safe: false })),
                    args: vec![],
                    optional: false })
            } else {
                sci
            }
        }
        Some('%') => {
            let scaled = call2("__pymul__", value, int_lit(100));
            let fixed = call2("__py_fmt_fixed", scaled, int_lit(precision.unwrap_or(6)));
            call2("__pyadd__", fixed, str_lit("%"))
        }
        None | Some('d') if grouping.is_some() && precision.is_none() => {
            let s = call_builtin1("str", value);
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__py_fmt_group".into()))),
                args: vec![Argument::positional(s)],
                optional: false })
        }
        None | Some('s') if grouping.is_none() && precision.is_none() && sign.is_none() => {
            call_builtin1("str", value)
        }
        _ => return None, // plain d/f/x/o/b/g etc → printf
    };
    if grouping.is_some() && !matches!(ty, None | Some('d')) {
        return None; // fractional/typed grouping not owned yet
    }

    // Alignment / width. Explicit align wins; a bare width right-aligns these
    // (numeric) values, matching Python's default for numbers.
    let apply = |base: Expression, method: &str, w: i64, fill: char| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(base),
                field: method.into(),
                null_safe: false })),
            // Fill is ALWAYS explicit: rjust/ljust (str_pad_*) drop content on a
            // defaulted fill.
            args: vec![
                Argument::positional(int_lit(w)),
                Argument::positional(str_lit(&fill.to_string())),
            ],
            optional: false })
    };
    match (align, width) {
        (Some(a), Some(w)) => {
            let method = match a {
                '<' => "ljust",
                '>' => "rjust",
                '^' => "center",
                _ => return None };
            Some(apply(base, method, w, fill))
        }
        (Some(_), None) => Some(base), // align without width is a no-op
        (None, Some(w)) if numeric_base => {
            // Numeric base with a bare width: right-align, zero-fill if requested.
            Some(apply(base, "rjust", w, if zero { '0' } else { ' ' }))
        }
        (None, Some(_)) => None, // string base + bare width → printf (runtime default align)
        (None, None) => Some(base) }
}

fn walk_fstring(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut parts = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::fstring_start | Rule::fstring_end => {}
            Rule::fstring_text => {
                // Literal segments carry backslash escapes (`\n`, `\t`, …) just
                // like a normal string literal, but the fstring grammar captures
                // them raw — process them here (same lowering as
                // `parse_python_string`).
                let unescaped = p
                    .as_str()
                    .replace("\\n", "\n")
                    .replace("\\t", "\t")
                    .replace("\\r", "\r")
                    .replace("\\\\", "\\")
                    .replace("\\'", "'")
                    .replace("\\\"", "\"");
                parts.push(InterpolPart::Text(unescaped));
            }
            Rule::fstring_escaped_brace => {
                let text = if p.as_str().starts_with('{') {
                    "{"
                } else {
                    "}"
                };
                parts.push(InterpolPart::Text(text.into()));
            }
            Rule::fstring_expr => {
                // `{ expr [!conv] [:spec] }` — mirror the `.format()` path
                // (`expand_format_field`): apply the conversion, then the format
                // spec, and otherwise wrap in `str()` so bool/None/containers
                // render with Python semantics (`True`/`None`/`[1, 2]`) instead
                // of the generic JS `toString` (`true`/`null`/`1,2`).
                let mut base: Option<Expression> = None;
                let mut conv: Option<char> = None;
                let mut spec: Option<String> = None;
                // `{expr=}` — the self-documenting form echoes the expression's
                // source text plus `=` before the value.
                let mut debug_src: Option<String> = None;
                let mut is_debug = false;
                for fp in p.into_inner() {
                    match fp.as_rule() {
                        Rule::fstring_debug => is_debug = true,
                        Rule::fstring_conversion => {
                            conv = fp.as_str().trim_start_matches('!').chars().next();
                        }
                        Rule::fstring_spec => {
                            spec = fp
                                .into_inner()
                                .find(|s| s.as_rule() == Rule::fstring_format_spec)
                                .map(|s| s.as_str().to_string());
                        }
                        Rule::fstring_format_spec => {
                            spec = Some(fp.as_str().to_string());
                        }
                        r if is_expression_rule(r) => {
                            debug_src = Some(fp.as_str().to_string());
                            base = Some(walk_expression(fp)?);
                        }
                        _ => {}
                    }
                }
                let Some(base) = base.map(desugar_member_reads) else {
                    continue;
                };
                // Emit `<source>=` as literal text; the value then renders with
                // `repr` unless a conversion or format spec was given.
                if is_debug {
                    if let Some(src) = &debug_src {
                        parts.push(InterpolPart::Text(format!("{src}=")));
                    }
                    if conv.is_none() && spec.is_none() {
                        conv = Some('r');
                    }
                }

                // Conversion first (Python order: convert, then format).
                let converted = match conv {
                    Some('r') | Some('a') => call_builtin1("repr", base),
                    Some('s') => call_builtin1("str", base),
                    _ => base };

                // A supported spec formats via `fmt % value` (`__pymod__`); an
                // unsupported one (`^` centre, `%`, grouping) falls back to
                // `str()` so the field still renders Python-correctly minus the
                // padding, rather than dropping to raw JS coercion.
                let field = match spec.as_deref().filter(|s| !s.is_empty()) {
                    Some(spec) => {
                        if let Some(native) = expand_python_format_spec(converted.clone(), spec) {
                            native
                        } else if let Some(fmt) = format_spec_to_printf(spec) {
                            Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Ident(
                                    "__pymod__".into(),
                                ))),
                                args: vec![
                                    Argument::positional(Expression::new(ExprKind::Lit(
                                        Literal::Str(fmt),
                                    ))),
                                    Argument::positional(converted),
                                ],
                                optional: false })
                        } else {
                            call_builtin1("str", converted)
                        }
                    }
                    None if conv.is_none() => call_builtin1("str", converted),
                    None => converted };
                parts.push(InterpolPart::Expr(field));
            }
            _ => {}
        }
    }
    Ok(ExprKind::Interpolation(parts))
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

/// Convert an assignment-target expression element into a destructuring
/// pattern element, handling `*rest` (`ExprKind::Spread` → `Rest`) and nested
/// tuple/list targets (`(a, (b, c))` → nested `Array` pattern) — not just bare
/// identifiers.
fn expr_to_array_pattern_elem(e: &Expression) -> ArrayPatternElem {
    match &e.kind {
        ExprKind::Ident(name) => {
            ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
        }
        ExprKind::Spread(inner) => match &inner.kind {
            ExprKind::Ident(name) => ArrayPatternElem::Rest(name.clone()),
            _ => ArrayPatternElem::Hole },
        ExprKind::Tuple(elems) => {
            let nested = elems.iter().map(expr_to_array_pattern_elem).collect();
            ArrayPatternElem::Pattern(BindingPattern::Array(nested), None)
        }
        ExprKind::Array(elems) => {
            let nested = elems
                .iter()
                .map(|ae| expr_to_array_pattern_elem(&ae.value))
                .collect();
            ArrayPatternElem::Pattern(BindingPattern::Array(nested), None)
        }
        _ => ArrayPatternElem::Hole }
}

fn walk_expr_list(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    // A trailing comma makes a single element a 1-tuple (`x,` / `(x,)`), not a
    // scalar. pest consumes the comma silently, so recover it from the source.
    let trailing_comma = pair.as_str().trim_end().ends_with(',');
    let mut inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 && !trailing_comma {
        walk_expression(inner.remove(0))
    } else if inner.is_empty() {
        Ok(Expression::new(ExprKind::Tuple(Vec::new())))
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(Expression::with_span(ExprKind::Tuple(exprs), span))
    }
}

fn walk_expr_list_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // See `walk_expr_list`: a trailing comma on a single element is a 1-tuple.
    let trailing_comma = pair.as_str().trim_end().ends_with(',');
    let inner: Vec<Pair<Rule>> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .collect();
    if inner.len() == 1 && !trailing_comma {
        walk_expr_kind(inner.into_iter().next().unwrap())
    } else if inner.is_empty() {
        Ok(ExprKind::Tuple(Vec::new()))
    } else {
        let exprs = inner
            .into_iter()
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;
        Ok(ExprKind::Tuple(exprs))
    }
}

fn walk_expr_list_or_single(pair: Pair<Rule>) -> Result<Expression, String> {
    if pair.as_rule() == Rule::expression_list {
        walk_expr_list(pair)
    } else {
        walk_expression(pair)
    }
}

fn walk_remaining_as_expr(items: &mut Vec<Pair<Rule>>) -> Result<Expression, String> {
    if items.len() == 1 {
        walk_expression(items.remove(0))
    } else {
        walk_expression(items.remove(0))
    }
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (sl, sc) = s.start_pos().line_col();
    let (el, ec) = s.end_pos().line_col();
    Span {
        start_line: sl as u32,
        start_col: sc as u32,
        end_line: el as u32,
        end_col: ec as u32 }
}

fn next_meaningful<'a>(
    pairs: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        match p.as_rule() {
            Rule::NEWLINE
            | Rule::INDENT
            | Rule::DEDENT
            | Rule::in_kw
            | Rule::as_kw
            | Rule::async_kw => continue,
            _ => return Ok(p) }
    }
    Err("No more meaningful pairs".into())
}

fn next_rule_any<'a>(
    pairs: &mut impl Iterator<Item = Pair<'a, Rule>>,
    rules: &[Rule],
) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        if rules.contains(&p.as_rule()) {
            return Ok(p);
        }
    }
    Err(format!("Expected one of {:?}", rules))
}

fn is_expression_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::expression
            | Rule::expression_list
            | Rule::named_expr
            | Rule::ternary_expr
            | Rule::or_expr
            | Rule::and_expr
            | Rule::not_expr
            | Rule::comparison
            | Rule::bitor_expr
            | Rule::bitxor_expr
            | Rule::bitand_expr
            | Rule::shift_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::unary
            | Rule::power
            | Rule::await_expr
            | Rule::postfix
            | Rule::primary
            | Rule::lambda_expr
            | Rule::yield_expr
            | Rule::star_expr
            | Rule::fstring
            | Rule::numeric_literal
            | Rule::string_literal
            | Rule::string_concat
            | Rule::identifier
            | Rule::true_kw
            | Rule::false_kw
            | Rule::none_kw
            | Rule::ellipsis_lit
            | Rule::subscript
            | Rule::subscript_item
            | Rule::comp_for_arg
    )
}

fn is_op_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::additive_op
            | Rule::multiplicative_op
            | Rule::shift_op
            | Rule::comparison_op
            | Rule::aug_assign_op
    )
}

fn parse_number(s: &str) -> Result<ExprKind, String> {
    let s = s.replace('_', "");
    // Complex numbers (j suffix)
    if s.ends_with('j') || s.ends_with('J') {
        let num_str = &s[..s.len() - 1];
        let val: f64 = num_str.parse().unwrap_or(0.0);
        return Ok(ExprKind::Lit(Literal::Float(val)));
    }
    if s.contains('.')
        || (s.contains('e') || s.contains('E')) && !s.starts_with("0x") && !s.starts_with("0X")
    {
        Ok(ExprKind::Lit(Literal::Float(
            s.parse().map_err(|e| format!("{}", e))?,
        )))
    } else if s.starts_with("0x") || s.starts_with("0X") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 16).unwrap_or(0),
        )))
    } else if s.starts_with("0o") || s.starts_with("0O") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 8).unwrap_or(0),
        )))
    } else if s.starts_with("0b") || s.starts_with("0B") {
        Ok(ExprKind::Lit(Literal::Int(
            i64::from_str_radix(&s[2..], 2).unwrap_or(0),
        )))
    } else {
        Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
    }
}

fn is_bytes_prefix(s: &str) -> bool {
    let lc = s.to_ascii_lowercase();
    lc.starts_with("b'")
        || lc.starts_with("b\"")
        || lc.starts_with("rb'")
        || lc.starts_with("rb\"")
        || lc.starts_with("br'")
        || lc.starts_with("br\"")
}

fn parse_bytes_literal(s: &str) -> ExprKind {
    // `Literal::Bytes`, not a `__py_bytes_new__([…])` call: the AST node is what
    // gives the value a static type, so `b[0]` resolves the `(Bytes, GetItem)`
    // binding instead of sniffing the receiver's kind at runtime
    // (unifiedstringplan.md §3c, builtinslotplan.md §2c).
    ExprKind::Lit(Literal::Bytes(parse_python_bytes(s)))
}

/// Build a call to a named identifier with positional args.
fn call_ident(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false })
}

/// Construct a `bytes` value from an int-array expression as a real
/// `Uint8Array` (`ObjectKind::TypedArray`). The VM handles indexing and
/// iteration natively; display is detected at runtime via `arraybuffer.isView`
/// in `emit_py_repr`, so no static bytes-tracking is needed.
fn wrap_bytes(array: Expression) -> Expression {
    call_ident("__py_bytes_new__", vec![array])
}

fn positional_or_named_arg(args: &[Argument], index: usize, name: &str) -> Option<Expression> {
    args.iter()
        .find(|a| a.name.as_deref() == Some(name))
        .map(|a| a.value.clone())
        .or_else(|| args.get(index).map(|a| a.value.clone()))
}

fn python_int_byteorder_arg(args: &[Argument], index: usize) -> Expression {
    positional_or_named_arg(args, index, "byteorder").unwrap_or_else(|| Expression::string("big"))
}

fn parse_python_hex_float(s: &str) -> Option<f64> {
    let trimmed = s.trim();
    let (mantissa, exponent) = trimmed
        .strip_prefix("0x")
        .or_else(|| trimmed.strip_prefix("0X"))?
        .split_once(['p', 'P'])?;
    let exp: i32 = exponent.parse().ok()?;
    let (whole, frac) = mantissa.split_once('.').unwrap_or((mantissa, ""));
    let mut value = i64::from_str_radix(if whole.is_empty() { "0" } else { whole }, 16).ok()? as f64;
    let mut place = 16.0;
    for ch in frac.chars() {
        let digit = ch.to_digit(16)? as f64;
        value += digit / place;
        place *= 16.0;
    }
    Some(value * 2f64.powi(exp))
}

fn try_rewrite_python_numeric_method(
    object: &Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    match field {
        "bit_length" if args.is_empty() => {
            Some(call_ident("__py_int_bit_length__", vec![object.clone()]))
        }
        "bit_count" if args.is_empty() => {
            Some(call_ident("__py_int_bit_count__", vec![object.clone()]))
        }
        "is_integer" if args.is_empty() => {
            Some(call_ident("__py_num_is_integer__", vec![object.clone()]))
        }
        "as_integer_ratio" if args.is_empty() => {
            Some(call_ident("__py_float_as_integer_ratio__", vec![object.clone()]))
        }
        "conjugate" if args.is_empty() => Some(object.clone()),
        "to_bytes" => {
            let length = positional_or_named_arg(args, 0, "length")?;
            let byteorder = python_int_byteorder_arg(args, 1);
            Some(call_ident(
                "__py_int_to_bytes__",
                vec![object.clone(), length, byteorder],
            ))
        }
        "from_bytes"
            if matches!(&object.kind, ExprKind::Ident(n) if n == "int")
                && !args.is_empty() =>
        {
            let bytes = args[0].value.clone();
            let byteorder = python_int_byteorder_arg(args, 1);
            Some(call_ident("__py_int_from_bytes__", vec![bytes, byteorder]))
        }
        "fromhex"
            if matches!(&object.kind, ExprKind::Ident(n) if n == "float")
                && args.len() == 1 =>
        {
            if let ExprKind::Lit(Literal::Str(s)) = &args[0].value.kind {
                parse_python_hex_float(s).map(Expression::float)
            } else {
                None
            }
        }
        _ => None }
}

/// bytes methods that return `bytes` (re-encoded after the string op).
const BYTES_METHODS_RETURN_BYTES: &[&str] = &[
    "upper",
    "lower",
    "capitalize",
    "title",
    "swapcase",
    "replace",
    "strip",
    "lstrip",
    "rstrip",
    "center",
    "ljust",
    "rjust",
    "zfill",
];
/// bytes methods that return a scalar (int/bool) — no re-encode.
const BYTES_METHODS_RETURN_SCALAR: &[&str] = &[
    "find",
    "rfind",
    "count",
    "startswith",
    "endswith",
    "isalpha",
    "isdigit",
    "isalnum",
    "isspace",
];

/// True when `e` is statically known to evaluate to `bytes`.
fn expr_is_python_bytes(e: &Expression) -> bool {
    match &e.kind {
        // A `Literal::Bytes` is bytes by construction — the call shape below is
        // the pre-literal spelling, kept for `bytes(...)` conversions.
        ExprKind::Lit(Literal::Bytes(_)) => true,
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(n) if n == "__py_bytes_new__" || n == "bytes" => true,
            // `+`/`*` lower to __pyadd__/__pymul__ — bytes if an operand is.
            ExprKind::Ident(n) if n == "__pyadd__" || n == "__pymul__" => {
                args.iter().any(|a| expr_is_python_bytes(&a.value))
            }
            ExprKind::Member { object, field, .. } => {
                field == "encode"
                    || (BYTES_METHODS_RETURN_BYTES.contains(&field.as_str())
                        && expr_is_python_bytes(object))
            }
            _ => false },
        _ => false }
}

fn py_static_bytes_has_non_ascii(e: &Expression) -> bool {
    if let ExprKind::Lit(Literal::Bytes(bytes)) = &e.kind {
        return bytes.iter().any(|b| *b > 0x7f);
    }
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return false;
    };
    if !matches!(&callee.kind, ExprKind::Ident(n) if n == "__py_bytes_new__") || args.len() != 1 {
        return false;
    }
    let ExprKind::Array(elems) = &args[0].value.kind else {
        return false;
    };
    elems.iter().any(|elem| {
        matches!(&elem.value.kind, ExprKind::Lit(Literal::Int(n)) if *n > 0x7f)
    })
}

/// Decode a bytes argument (e.g. the needle of `find(b'x')`) to a latin-1
/// string; leave non-bytes args (widths, fill counts) untouched.
fn decode_bytes_arg(a: &Argument) -> Argument {
    if expr_is_python_bytes(&a.value) {
        Argument {
            value: call_ident("__vybe_bytes_decode", vec![a.value.clone()]),
            name: a.name.clone(),
            by_ref: a.by_ref,
            spread: a.spread }
    } else {
        a.clone()
    }
}

/// Rewrite a string-like method call on a `bytes` receiver as
/// decode → `str.METHOD(...)` → (re-encode if it returns bytes). Returns
/// `None` when the receiver isn't statically bytes or the method isn't a
/// supported string-like bytes method.
fn try_rewrite_bytes_method(
    object: &Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    if !expr_is_python_bytes(object) {
        return None;
    }
    // `.hex()` → uint8array.toHex (a hex string, no `0x`/separators).
    if field == "hex" && args.is_empty() {
        return Some(call_ident("__py_bytes_hex__", vec![object.clone()]));
    }
    // `.split()`/`.join()` are deferred: they involve a *list of bytes*, and
    // nested bytes don't yet repr as `b'…'` inside a list/collection.
    // `.decode(...)` is handled by the profile method entry (UTF-8) — leave it.
    let returns_bytes = BYTES_METHODS_RETURN_BYTES.contains(&field);
    let returns_scalar = BYTES_METHODS_RETURN_SCALAR.contains(&field);
    if !returns_bytes && !returns_scalar {
        return None;
    }
    // decode(receiver).METHOD(decode(arg0), …)
    let decoded_recv = call_ident("__vybe_bytes_decode", vec![object.clone()]);
    let decoded_args: Vec<Argument> = args.iter().map(decode_bytes_arg).collect();
    let str_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(decoded_recv),
            field: field.into(),
            null_safe: false })),
        args: decoded_args,
        optional: false });
    if returns_bytes {
        // re-encode the resulting str back to bytes
        Some(wrap_bytes(call_ident("__vybe_str_encode", vec![str_call])))
    } else {
        Some(str_call)
    }
}

fn parse_python_string(s: &str) -> String {
    let raw = s.starts_with('r')
        || s.starts_with('R')
        || s.starts_with("rb")
        || s.starts_with("rB")
        || s.starts_with("Rb")
        || s.starts_with("RB")
        || s.starts_with("br")
        || s.starts_with("bR")
        || s.starts_with("Br")
        || s.starts_with("BR");
    let mut s = s;
    // Strip prefix (r, b, rb, u, etc.)
    let prefixes = [
        "rb", "Rb", "rB", "RB", "br", "bR", "Br", "BR", "r", "R", "b", "B", "u", "U",
    ];
    for prefix in &prefixes {
        if s.starts_with(prefix) {
            s = &s[prefix.len()..];
            break;
        }
    }
    // Strip quotes
    if s.starts_with("\"\"\"") {
        s = &s[3..s.len() - 3];
    } else if s.starts_with("'''") {
        s = &s[3..s.len() - 3];
    } else if s.starts_with('"') {
        s = &s[1..s.len() - 1];
    } else if s.starts_with('\'') {
        s = &s[1..s.len() - 1];
    }
    if raw {
        return s.to_string();
    }
    decode_python_escape_str(s)
}

/// Unescape a `str` literal at the CHARACTER level.
///
/// The byte-level decoder is for `b"…"` literals. Running a `str` through it
/// and then `char::from(u8)` maps each byte to its LATIN-1 code point, so the
/// two UTF-8 bytes of `é` become `Ã` + `©` — every non-ASCII literal came out
/// mojibake, `len("héllo")` was 6, and `.encode("utf-8")` double-encoded.
///
/// Escapes that name a code point (`\xNN`, `\uNNNN`, `\UNNNNNNNN`) produce that
/// CHARACTER, per CPython: `"\xe9"` is `é` (U+00E9), not the byte 0xE9.
fn decode_python_escape_str(s: &str) -> String {
    let chars: Vec<char> = s.chars().collect();
    let mut out = String::with_capacity(s.len());
    let mut i = 0;
    // Read `n` hex digits starting at `from`, if all present and valid.
    let hex_at = |from: usize, n: usize| -> Option<u32> {
        if from + n > chars.len() {
            return None;
        }
        let mut v: u32 = 0;
        for c in &chars[from..from + n] {
            v = v * 16 + c.to_digit(16)?;
        }
        Some(v)
    };
    while i < chars.len() {
        if chars[i] != '\\' || i + 1 >= chars.len() {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        match chars[i + 1] {
            'n' => {
                out.push('\n');
                i += 2;
            }
            't' => {
                out.push('\t');
                i += 2;
            }
            'r' => {
                out.push('\r');
                i += 2;
            }
            '0' => {
                out.push('\0');
                i += 2;
            }
            'a' => {
                out.push('\u{7}');
                i += 2;
            }
            'b' => {
                out.push('\u{8}');
                i += 2;
            }
            'f' => {
                out.push('\u{c}');
                i += 2;
            }
            'v' => {
                out.push('\u{b}');
                i += 2;
            }
            '\\' | '\'' | '"' => {
                out.push(chars[i + 1]);
                i += 2;
            }
            // Line continuation: a backslash before a newline emits nothing.
            '\n' => {
                i += 2;
            }
            'x' | 'u' | 'U' => {
                let n = match chars[i + 1] {
                    'x' => 2,
                    'u' => 4,
                    _ => 8 };
                match hex_at(i + 2, n).and_then(char::from_u32) {
                    Some(c) => {
                        out.push(c);
                        i += 2 + n;
                    }
                    None => {
                        out.push(chars[i + 1]);
                        i += 2;
                    }
                }
            }
            other => {
                // Unknown escape — CPython keeps the backslash AND the char.
                out.push('\\');
                out.push(other);
                i += 2;
            }
        }
    }
    out
}

fn parse_python_bytes(s: &str) -> Vec<u8> {
    let raw = s.starts_with('r')
        || s.starts_with('R')
        || s.starts_with("rb")
        || s.starts_with("rB")
        || s.starts_with("Rb")
        || s.starts_with("RB")
        || s.starts_with("br")
        || s.starts_with("bR")
        || s.starts_with("Br")
        || s.starts_with("BR");
    let mut s = s;
    let prefixes = [
        "rb", "Rb", "rB", "RB", "br", "bR", "Br", "BR", "r", "R", "b", "B", "u", "U",
    ];
    for prefix in &prefixes {
        if s.starts_with(prefix) {
            s = &s[prefix.len()..];
            break;
        }
    }
    if s.starts_with("\"\"\"") {
        s = &s[3..s.len() - 3];
    } else if s.starts_with("'''") {
        s = &s[3..s.len() - 3];
    } else if s.starts_with('"') {
        s = &s[1..s.len() - 1];
    } else if s.starts_with('\'') {
        s = &s[1..s.len() - 1];
    }
    if raw {
        s.as_bytes().to_vec()
    } else {
        decode_python_escape_bytes(s)
    }
}

fn decode_python_escape_bytes(s: &str) -> Vec<u8> {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'\\' || i + 1 >= bytes.len() {
            out.push(bytes[i]);
            i += 1;
            continue;
        }
        match bytes[i + 1] {
            b'n' => {
                out.push(b'\n');
                i += 2;
            }
            b't' => {
                out.push(b'\t');
                i += 2;
            }
            b'r' => {
                out.push(b'\r');
                i += 2;
            }
            b'\\' | b'\'' | b'"' => {
                out.push(bytes[i + 1]);
                i += 2;
            }
            b'x' if i + 3 < bytes.len() => {
                if let (Some(hi), Some(lo)) = (hex_value(bytes[i + 2]), hex_value(bytes[i + 3])) {
                    out.push((hi << 4) | lo);
                    i += 4;
                } else {
                    out.push(bytes[i + 1]);
                    i += 2;
                }
            }
            other => {
                out.push(other);
                i += 2;
            }
        }
    }
    out
}

fn hex_value(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(b - b'a' + 10),
        b'A'..=b'F' => Some(b - b'A' + 10),
        _ => None }
}

fn parse_comparison_op(s: &str) -> BinOp {
    match s {
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "in" => BinOp::In,
        "is" => BinOp::Is,
        _ if s.contains("not") && s.contains("in") => BinOp::NotIn,
        _ if s.contains("is") && s.contains("not") => BinOp::IsNot,
        _ => BinOp::Eq }
}

fn parse_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "//" => BinOp::FloorDiv,
        "%" => BinOp::Mod,
        "**" => BinOp::Pow,
        "@" => BinOp::MatMul,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "&" => BinOp::BitAnd,
        _ => BinOp::Add }
}

fn parse_literal_to_expr(text: &str) -> Expression {
    if let Ok(n) = text.parse::<i64>() {
        Expression::int(n)
    } else if let Ok(f) = text.parse::<f64>() {
        Expression::float(f)
    } else {
        Expression::string(text)
    }
}
