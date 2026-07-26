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
                _ => break,
            }
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
    PY_SYS_MODULES_BOUND.with(|b| b.set(false));
    PY_DEFINED_CLASSES.with(|m| m.borrow_mut().clear());
    PY_DEFINED_FUNCTIONS.with(|m| m.borrow_mut().clear());
    PY_CALLABLE_CLASSES.with(|m| m.borrow_mut().clear());
    PY_CLASS_PARENTS.with(|m| m.borrow_mut().clear());
    PY_CLASS_ATTRS.with(|m| m.borrow_mut().clear());
    PY_NAMEDTUPLE_DEFS.with(|m| m.borrow_mut().clear());
    PY_NAMEDTUPLE_INSTANCES.with(|m| m.borrow_mut().clear());
    PY_SQL_VARS.with(|m| m.borrow_mut().clear());
    PY_RE_VARS.with(|m| m.borrow_mut().clear());
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
                _ => walk_stmt_into(pair, &mut body, &mut imports)?,
            }
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

    Ok(Module {
        name: "main".into(),
        language: Lang::Python,
        body,
        imports,
    })
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
fn parse_python_prelude(src: &str) -> Vec<Statement> {
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

/// `fnmatch` — shell-style pattern matching. Self-contained iterative matcher
/// (`*`, `?`) over local strings; `translate` builds a regex-shaped string.
const FNMATCH_PRELUDE: &str = r#"
def __fn_match(name, pat):
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
        return __fn_match(name.lower(), pat.lower())
    def fnmatchcase(self, name, pat):
        return __fn_match(name, pat)
    def filter(self, names, pat):
        result = []
        for nm in names:
            if __fn_match(nm.lower(), pat.lower()):
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
                if !py_known_module(&root) {
                    body.push(py_import_error_stmt(&format!("No module named '{root}'")));
                    return Ok(());
                }
                body.extend(py_module_rename_stmts(&root));
                let bound = alias.clone().unwrap_or_else(|| root.clone());
                if root == "sys" && !PY_SYS_MODULES_BOUND.with(|b| b.get()) {
                    PY_SYS_MODULES_BOUND.with(|b| b.set(true));
                    let props: Vec<ObjectProperty> = PY_IMPORTED_MODULES.with(|m| {
                        m.borrow()
                            .iter()
                            .map(|name| ObjectProperty::KeyValue {
                                key: Expression::new(ExprKind::Lit(Literal::Str(
                                    name.clone().into(),
                                ))),
                                value: Expression::new(ExprKind::Ident(name.clone())),
                            })
                            .collect()
                    });
                    body.push(Statement::new(StmtKind::Assign {
                        targets: vec![Expression::new(ExprKind::Ident("__py_sys_modules".into()))],
                        value: Expression::new(ExprKind::Object(props)),
                    }));
                } else if PY_SYS_MODULES_BOUND.with(|b| b.get()) {
                    body.push(Statement::new(StmtKind::Assign {
                        targets: vec![Expression::new(ExprKind::Index {
                            object: Box::new(Expression::new(ExprKind::Ident(
                                "__py_sys_modules".into(),
                            ))),
                            index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                                bound.clone().into(),
                            )))),
                            null_safe: false,
                        })],
                        value: Expression::new(ExprKind::Ident(bound)),
                    }));
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
                    // `deque` IS a JS array under the hood (same as list) —
                    // bind the ctor to a list copy; append/pop are already
                    // array emits, appendleft/popleft map to unshift/shift.
                    for n in names {
                        if n.name == "deque" {
                            let local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            body.push(Statement::new(StmtKind::Assign {
                                targets: vec![Expression::new(ExprKind::Ident(local))],
                                value: Expression::new(ExprKind::Lambda {
                                    params: vec![lambda_param("__it")],
                                    body: LambdaBody::Expr(Box::new(Expression::new(
                                        ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Ident(
                                                "list".into(),
                                            ))),
                                            args: vec![Argument::positional(Expression::new(
                                                ExprKind::Ident("__it".into()),
                                            ))],
                                            optional: false,
                                        },
                                    ))),
                                    is_async: false,
                                    captures: vec![],
                                }),
                            }));
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
                                value: lambda,
                            }));
                        }
                    }
                    return Ok(());
                }
            }
            imports.push(import);
        }
        _ => body.push(walk_statement(pair)?),
    }
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
                right: Box::new(Expression::new(ExprKind::Ident("__b".into()))),
            }))),
            is_async: false,
            captures: vec![],
        })
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
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        })
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
                expr: Box::new(Expression::new(ExprKind::Ident("__a".into()))),
            }))),
            is_async: false,
            captures: vec![],
        }),
        _ => return None,
    })
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
        is_nullable: false,
    }
}

fn py_builtin_callable_lambda(name: &str) -> Option<Expression> {
    let param = "__py_key_value";
    Some(match name {
        "len" => Expression::new(ExprKind::Lambda {
            params: vec![lambda_param(param)],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("len".into()))),
                args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                    param.into(),
                )))],
                optional: false,
            }))),
            is_async: false,
            captures: vec![],
        }),
        _ => return None,
    })
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

        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
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
        _ => false,
    }
}

fn expr_has_yield(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
        ExprKind::Call { args, .. } => args.iter().any(|a| expr_has_yield(&a.value)),
        ExprKind::Binary { left, right, .. } => expr_has_yield(left) || expr_has_yield(right),
        ExprKind::Unary { expr: e, .. } => expr_has_yield(e),
        ExprKind::Index { object, index, .. } => expr_has_yield(object) || expr_has_yield(index),
        _ => false,
    }
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
        is_sub: false,
    })
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
                            is_nullable: false,
                        });
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
                            is_nullable: false,
                        });
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
                            is_nullable: false,
                        });
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

fn other_attr(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::Ident("other".into()))),
        field: field.to_string(),
        null_safe: false,
    })
}

fn binop(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

/// `__repr__` as `@dataclass` generates it: `ClassName(field=repr, ...)`.
fn dataclass_repr(class_name: &str, fields: &[(String, Option<Expression>)]) -> Statement {
    let mut expr = str_lit(&format!("{class_name}("));
    for (i, (name, _)) in fields.iter().enumerate() {
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
fn dataclass_eq(class_name: &str, fields: &[(String, Option<Expression>)]) -> Statement {
    let same_class = binop(
        BinOp::Eq,
        call_ident("type", vec![Expression::new(ExprKind::Ident("other".into()))]),
        Expression::new(ExprKind::Ident(class_name.to_string())),
    );
    let mut cond = same_class;
    for (name, _) in fields {
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
        _ => false,
    }
}

/// The annotated class-level declarations, in source order — exactly what
/// CPython's `@dataclass` treats as fields. A bare `x = 5` with no annotation
/// is NOT a field, which is why the type hint (threaded through
/// `VarDeclarator.type_hint`) is the marker.
fn dataclass_fields(body: &[Statement]) -> Vec<(String, Option<Expression>)> {
    let mut fields = Vec::new();
    for stmt in body {
        let StmtKind::VarDecl { declarations, .. } = &stmt.kind else {
            continue;
        };
        for d in declarations {
            if d.type_hint.is_none() {
                continue;
            }
            if let BindingPattern::Ident(name) = &d.pattern {
                fields.push((name.clone(), d.init.clone()));
            }
        }
    }
    fields
}

fn self_attr(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::Ident("self".into()))),
        field: field.to_string(),
        null_safe: false,
    })
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
        is_nullable: false,
    }
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
        is_sub: false,
    })
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
        for (name, default) in &fields {
            params.push(plain_param(name, default.clone()));
            init_body.push(Statement::new(StmtKind::Assign {
                targets: vec![self_attr(name)],
                value: Expression::new(ExprKind::Ident(name.clone())),
            }));
        }
        body.push(fn_decl("__init__", params, init_body));
    }
    if !has_method(body, "__repr__") {
        body.push(dataclass_repr(class_name, &fields));
    }
    if !has_method(body, "__eq__") {
        body.push(dataclass_eq(class_name, &fields));
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
    let mut attrs = std::collections::HashSet::new();
    for stmt in &body_stmts {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, .. } => {
                attrs.insert(name.clone());
            }
            StmtKind::Assign { targets, .. } => {
                for target in targets {
                    if let ExprKind::Ident(attr) = &target.kind {
                        attrs.insert(attr.clone());
                    }
                }
            }
            _ => {}
        }
    }
    note_class_attrs(&name, attrs);
    if has_call_method {
        note_callable_class(&name);
    }

    // `@dataclass` — synthesize the members CPython's decorator generates at
    // runtime. Done here, in the walker, because decorators never reach
    // `normalize_class` (the shared signature carries modifiers, not
    // decorators) and because synthesizing real AST members keeps the shared
    // class pipeline language-neutral.
    if decorators.iter().any(is_dataclass_decorator) {
        synthesize_dataclass_members(&name, &mut body_stmts);
    }

    // Convert body statements into ClassMembers
    let members = stmts_to_class_members(body_stmts);

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    })
}

fn stmts_to_class_members(stmts: Vec<Statement>) -> Vec<ClassMember> {
    let mut members: Vec<ClassMember> = Vec::new();
    // Track Property member index by name so @x.setter can find the getter.
    let mut property_indices: std::collections::HashMap<String, usize> =
        std::collections::HashMap::new();

    for stmt in stmts {
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
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
                        visibility: Visibility::Public,
                    });
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
                        modifiers: Modifiers::default(),
                    });
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
                                is_nullable: false,
                            });
                            *setter = Some(vybe_ast::PropertySetter {
                                param: value_param,
                                body: body.clone(),
                            });
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
                        is_nullable: false,
                    };
                    let mut p = vec![dummy];
                    p.extend_from_slice(params);
                    p
                } else {
                    params.clone()
                };
                let is_static =
                    has_staticmethod || final_params.first().map_or(true, |p| p.name != "self");
                let mut mods = modifiers.clone();
                mods.is_static = is_static;
                members.push(ClassMember::Method(Box::new(Statement::new(
                    StmtKind::FunctionDecl {
                        name: name.clone(),
                        params: final_params,
                        return_type: None,
                        body: body.clone(),
                        modifiers: mods,
                        handles: Vec::new(),
                        is_async: *is_async,
                        is_generator: false,
                        is_sub: false,
                    },
                ))));
            }
            // Annotated class-level declaration (`x: int = 0`, or bare
            // `x: int`). The type hint is what marks it a dataclass field, so
            // it is carried onto the member rather than dropped.
            StmtKind::VarDecl { declarations, .. } => {
                for d in declarations {
                    let BindingPattern::Ident(field_name) = &d.pattern else {
                        continue;
                    };
                    let mut mods = Modifiers::default();
                    mods.is_static = true; // Python class-level vars are class attributes
                    members.push(ClassMember::Field {
                        name: field_name.clone(),
                        type_hint: d.type_hint.clone(),
                        init: d.init.clone(),
                        modifiers: mods,
                        with_events: false,
                        array_bounds: None,
                    });
                }
            }
            StmtKind::Assign { targets, value } => {
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
                            array_bounds: None,
                        });
                    }
                }
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
        _ => false,
    }
}

/// Desugar Python function decorators to runtime application:
/// `@a @b def f(...)` → `f = a(b(<function f>))`. Fires only when every
/// decorator is a general (user) decorator; if any is special the declaration
/// is returned unchanged so the specialized compile paths still see it.
fn desugar_function_decorators(decl: StmtKind, decorators: Vec<Expression>) -> StmtKind {
    if decorators.is_empty() || decorators.iter().any(is_special_decorator) {
        return decl;
    }
    let StmtKind::FunctionDecl { name, .. } = &decl else {
        return decl;
    };
    let fn_name = name.clone();
    // Strip the now-runtime-applied decorators off the inner declaration so the
    // metadata pass doesn't ALSO treat them as (inert) annotations.
    let mut inner = decl;
    if let StmtKind::FunctionDecl { modifiers, .. } = &mut inner {
        modifiers.decorators = Vec::new();
    }
    let func_expr = Expression::new(ExprKind::FunctionExpr(Box::new(Statement {
        kind: inner,
        span: Span::default(),
    })));
    // Fold innermost-first (reversed) so `@a @b def f` becomes `a(b(f))`.
    let mut acc = func_expr;
    for d in decorators.into_iter().rev() {
        acc = Expression::new(ExprKind::Call {
            callee: Box::new(d),
            args: vec![Argument {
                value: acc,
                name: None,
                by_ref: false,
                spread: false,
            }],
            optional: false,
        });
    }
    StmtKind::Assign {
        targets: vec![Expression::ident(&fn_name)],
        value: acc,
    }
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
            Rule::class_def => walk_class_def(item, decorators),
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
            )),
        }
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
        else_body,
    })
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
                    value: Expression::bool(true),
                }));
                out.push(stmt);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                out.push(Statement::new(StmtKind::If {
                    cond,
                    then_body: mark_loop_break_sets_flag(then_body, flag),
                    elifs: elifs
                        .into_iter()
                        .map(|(c, b)| (c, mark_loop_break_sets_flag(b, flag)))
                        .collect(),
                    else_body: else_body.map(|b| mark_loop_break_sets_flag(b, flag)),
                }));
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
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
                    finally: finally.map(|b| mark_loop_break_sets_flag(b, flag)),
                }));
            }
            StmtKind::With {
                items,
                body,
                is_async,
            } => {
                out.push(Statement::new(StmtKind::With {
                    items,
                    body: mark_loop_break_sets_flag(body, flag),
                    is_async,
                }));
            }
            StmtKind::Block(b) => {
                out.push(Statement::new(StmtKind::Block(mark_loop_break_sets_flag(
                    b, flag,
                ))));
            }
            _ => out.push(stmt),
        }
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
            else_body: None,
        });
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
        value: Expression::bool(false),
    });
    let while_stmt = Statement::new(StmtKind::While {
        cond,
        body,
        else_body: None,
    });
    let else_guard = Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(Expression::ident(&flag)),
        }),
        then_body: else_stmts,
        elifs: Vec::new(),
        else_body: None,
    });
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
                    null_safe: false,
                }),
            }));
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
        is_async,
    })
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
                            // Exception type expression
                            types.push(cp.as_str().trim().to_string());
                        }
                    }
                }
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
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

    Ok(StmtKind::Try {
        body,
        catches,
        else_body,
        finally,
    })
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
                    var,
                });
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
        null_safe: false,
    })
}
fn with_arg(value: Expression) -> Argument {
    Argument {
        value,
        name: None,
        by_ref: false,
        spread: false,
    }
}
fn with_call(callee: Expression, args: Vec<Argument>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}
fn with_not(e: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(e),
    })
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
                value: Expression::bool(true),
            }),
            sql_call("__sql_rollback"),
            with_stmt(StmtKind::Throw {
                expr: Some(Expression::ident(&exc)),
                cause: None,
            }),
        ],
        when_clause: None,
    };

    let finally = vec![with_stmt(StmtKind::If {
        cond: with_not(Expression::ident(&hit)),
        then_body: vec![sql_call("__sql_commit")],
        elifs: vec![],
        else_body: None,
    })];

    vec![
        sql_call("__sql_begin"),
        with_stmt(StmtKind::Assign {
            targets: vec![Expression::ident(&hit)],
            value: Expression::bool(false),
        }),
        with_stmt(StmtKind::Try {
            body,
            catches: vec![catch],
            else_body: None,
            finally: Some(finally),
        }),
    ]
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
            value,
        })
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
                    cause: None,
                })],
                elifs: vec![],
                else_body: None,
            }),
        ],
        when_clause: None,
    };

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
        else_body: None,
    })];

    vec![
        assign(&mgr, first.expr.clone()),
        assign(&target, with_call(with_member(&mgr, "__enter__"), vec![])),
        assign(&hit, Expression::bool(false)),
        with_stmt(StmtKind::Try {
            body: inner_body,
            catches: vec![catch],
            else_body: None,
            finally: Some(finally),
        }),
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
                    body,
                });
            }
            Rule::NEWLINE | Rule::INDENT | Rule::DEDENT => {}
            _ => {}
        }
    }

    Ok(StmtKind::MatchStatement {
        subject: subject.unwrap_or(Expression::new(ExprKind::Lit(Literal::Null))),
        cases,
    })
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
                None => Ok(Pattern::Wildcard),
            }
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
                name,
            })
        }
        Rule::wildcard_pattern => Ok(Pattern::Wildcard),
        Rule::capture_pattern => {
            let name = pair.as_str().to_string();
            Ok(Pattern::As {
                pattern: None,
                name: Some(name),
            })
        }
        Rule::singleton_pattern => {
            let text = pair.as_str().trim();
            let expr = match text {
                "None" => Expression::null(),
                "True" => Expression::bool(true),
                "False" => Expression::bool(false),
                _ => Expression::null(),
            };
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
                    _ => patterns.push(walk_pattern(cp)?),
                }
            }
            Ok(Pattern::Class {
                cls: Expression::new(ExprKind::Ident(cls_name)),
                patterns,
                kw_patterns,
            })
        }
        Rule::true_kw => Ok(Pattern::Singleton(Expression::bool(true))),
        Rule::false_kw => Ok(Pattern::Singleton(Expression::bool(false))),
        Rule::none_kw => Ok(Pattern::Singleton(Expression::null())),
        other => Err(format!("Unexpected pattern rule: {:?}", other)),
    }
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
                if saw_from {
                    cause = Some(walk_expression(p)?);
                } else {
                    exc = Some(walk_expression(p)?);
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
                        null_safe: false,
                    })),
                    args: vec![Argument::positional((**index).clone())],
                    optional: false,
                });
                return Ok(StmtKind::Expr(pop));
            }
        }
    }
    Ok(StmtKind::Delete(exprs))
}

fn walk_assert(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs: Vec<Expression> = pair
        .into_inner()
        .filter(|p| is_expression_rule(p.as_rule()))
        .map(walk_expression)
        .collect::<Result<Vec<_>, _>>()?;
    let msg = if exprs.len() > 1 { exprs.pop() } else { None };
    let test = exprs.into_iter().next().unwrap_or(Expression::bool(false));
    Ok(StmtKind::Assert { test, msg })
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
        let target = walk_expr_list_or_single(inner.remove(0))?;
        let op_str = inner.remove(0).as_str(); // aug_assign_op
        let value = if inner.len() == 1 {
            walk_expr_list_or_single(inner.remove(0))?
        } else {
            walk_remaining_as_expr(&mut inner)?
        };
        // `+=` / `*=` use Python's dynamic add/mul (list concat/repeat, string
        // ops), so lower to `target = __pyadd__(target, value)` — the numeric
        // CompoundAssign path coerces operands to f64 and traps on lists.
        if op_str == "+=" || op_str == "*=" {
            let helper = if op_str == "+=" {
                "__pyadd__"
            } else {
                "__pymul__"
            };
            let combined = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(helper.into()))),
                args: vec![
                    Argument::positional(target.clone()),
                    Argument::positional(value),
                ],
                optional: false,
            });
            return Ok(StmtKind::Assign {
                targets: vec![target],
                value: combined,
            });
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
            _ => None,
        };
        if let Some(op) = set_binop {
            let combined = Expression::new(ExprKind::Binary {
                op,
                left: Box::new(target.clone()),
                right: Box::new(value),
            });
            return Ok(StmtKind::Assign {
                targets: vec![target],
                value: combined,
            });
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
            _ => CompoundOp::Add,
        };
        return Ok(StmtKind::CompoundAssign { target, op, value });
    }

    // Check if this has "=" tokens — simple assignment
    // The grammar captures: expression_list ~ ("=" ~ expression_list)+
    // So we may have multiple expression_list separated by = signs
    if inner.len() == 1 {
        let expr = walk_expr_list_or_single(inner.remove(0))?;
        return Ok(StmtKind::Expr(expr));
    }

    // Multiple items => chained assignment (`a = b = c`) or an annotated
    // assignment (`x: int = val`). `type_annotation` is skipped rather than
    // collected: it is a TYPE, not an assignment target. Treating it as one
    // made `x: int = 5` compile as `x = int = 5`, rebinding the name `int`.
    let mut annotation: Option<String> = None;
    let mut all_exprs = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::type_annotation {
            annotation = Some(p.as_str().trim().to_string());
            continue;
        }
        if is_expression_rule(p.as_rule()) || p.as_rule() == Rule::expression_list {
            all_exprs.push(walk_expr_list_or_single(p)?);
        }
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
                    value,
                },
                None => StmtKind::Empty,
            });
        };
        return Ok(StmtKind::VarDecl {
            declarations: vec![vybe_ast::VarDeclarator {
                pattern: BindingPattern::Ident(name.clone()),
                type_hint: Some(hint),
                init,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        });
    }

    if all_exprs.len() >= 2 {
        let mut value = all_exprs.pop().unwrap();
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
                            defaults: Vec::new(),
                        },
                    );
                }
            }
        }
        // `m = json` / `m = importlib.import_module('json')` (the walker
        // already lowered the latter to the module Ident): record the
        // local as a module alias so member access substitutes the root.
        if all_exprs.len() == 1 {
            if let ExprKind::Ident(target_name) = &all_exprs[0].kind {
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
            }
        }
        // Track sqlite handles: `conn = sqlite3.connect(...)` / `cur = conn.cursor()`
        // so later `.execute()`/`.close()` on them route to the `__sql_*` builtins.
        for t in &all_exprs {
            note_sql_var_if_producer(t, &value);
        }
        // Convert Tuple targets to Destructure for tuple unpacking (x, y = ...)
        let targets = all_exprs
            .into_iter()
            .map(|t| {
                if let ExprKind::Tuple(elems) = &t.kind {
                    let patterns = elems.iter().map(expr_to_array_pattern_elem).collect();
                    Expression::new(ExprKind::Destructure(DestructurePattern::Array(patterns)))
                } else {
                    t
                }
            })
            .collect();
        Ok(StmtKind::Assign { targets, value })
    } else if all_exprs.len() == 1 {
        Ok(StmtKind::Expr(all_exprs.remove(0)))
    } else {
        Ok(StmtKind::Empty)
    }
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
            span,
        })
    } else {
        // Multiple: import os, sys — emit first, rest are separate
        let (path, alias) = imports.remove(0);
        Ok(Import {
            kind: ImportKind::Simple { path, alias },
            span,
        })
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

    if is_wildcard {
        Ok(Import {
            kind: ImportKind::Wildcard {
                path: module,
                alias: None,
            },
            span,
        })
    } else {
        Ok(Import {
            kind: ImportKind::Named {
                path: module,
                names,
                level,
            },
            span,
        })
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

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
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
                    value: Box::new(value),
                })
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
                    else_: Box::new(orelse),
                })
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
            // The conditional's own condition already applies Python truthiness
            // (`if []:` is falsy), so use the operand directly as the condition.
            Ok(ExprKind::Ternary {
                cond: Box::new(operand),
                then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                else_: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
            })
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
                                null_safe: false,
                            })),
                            args: vec![Argument::positional(left.clone())],
                            optional: false,
                        });
                        left = if op == BinOp::NotIn {
                            Expression::new(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(has_call),
                            })
                        } else {
                            has_call
                        };
                    } else if matches!(op, BinOp::In | BinOp::NotIn) {
                        // `x in y` — polymorphic membership (string substring /
                        // list element / dict key). Route to the Python adapter
                        // `__py_contains__(y, x)` rather than the shared
                        // `BinOp::In`, whose runtime array-classification
                        // mis-sends plain objects to `Array.includes`.
                        let contains = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(
                                "__py_contains__".into(),
                            ))),
                            args: vec![
                                Argument::positional(right),
                                Argument::positional(left.clone()),
                            ],
                            optional: false,
                        });
                        left = if op == BinOp::NotIn {
                            Expression::new(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(contains),
                            })
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
                                optional: false,
                            });
                        }
                    } else if matches!(op, BinOp::Eq | BinOp::NotEq) {
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
                                right: Box::new(call_ident("__vybe_bytes_decode", vec![right])),
                            });
                        } else {
                            left = Expression::new(ExprKind::Binary {
                                op,
                                left: Box::new(left.clone()),
                                right: Box::new(right),
                            });
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
                            right: Box::new(call_ident("__vybe_bytes_decode", vec![right])),
                        });
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
                            optional: false,
                        });
                    } else {
                        left = Expression::new(ExprKind::Binary {
                            op,
                            left: Box::new(left),
                            right: Box::new(right),
                        });
                    }
                }
                Ok(left.kind)
            } else {
                let mut result = Expression::new(ExprKind::Binary {
                    op: comparisons[0].0,
                    left: Box::new(operands[0].clone()),
                    right: Box::new(operands[1].clone()),
                });
                for j in 1..comparisons.len() {
                    let pairwise = Expression::new(ExprKind::Binary {
                        op: comparisons[j].0,
                        left: Box::new(operands[j].clone()),
                        right: Box::new(operands[j + 1].clone()),
                    });
                    result = Expression::new(ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(result),
                        right: Box::new(pairwise),
                    });
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
                    optional: false,
                });
            }
            let op = match op_str {
                "-" => UnaryOp::Neg,
                "+" => UnaryOp::Pos,
                "~" => UnaryOp::BitNot,
                _ => UnaryOp::Neg,
            };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(operand),
            })
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
                    optional: false,
                })
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
            left = Expression::new(ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
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
                    let helper = if op_str == "+" {
                        "__pyadd__"
                    } else {
                        "__pysub__"
                    };
                    left = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident(helper.into()))),
                        args: vec![Argument::positional(left), Argument::positional(right)],
                        optional: false,
                    });
                } else {
                    let op = parse_binop(op_str);
                    left = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right),
                    });
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
                    _ => None,
                };
                if let Some(helper) = helper {
                    let callee = Expression::new(ExprKind::Ident(helper.into()));
                    left = Expression::new(ExprKind::Call {
                        callee: Box::new(callee),
                        args: vec![Argument::positional(left), Argument::positional(right)],
                        optional: false,
                    });
                } else {
                    let op = parse_binop(op_str);
                    left = Expression::new(ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right),
                    });
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
                    right: Box::new(right),
                });
            }
        } else if is_expression_rule(p.as_rule()) {
            // Operator was merged into the rule text, parse from context
            let right = walk_expression(items[i].clone())?;
            i += 1;
            left = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(left),
                right: Box::new(right),
            });
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
            is_nullable: false,
        }],
        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(key),
            args: vec![Argument::positional(Expression::ident("__sk"))],
            optional: false,
        }))),
        is_async: false,
        captures: vec![],
    })
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
        is_nullable: false,
    };
    Expression::new(ExprKind::Lambda {
        params: vec![param("__cur"), param("__row")],
        body: LambdaBody::Expr(Box::new(Expression::ident("__row"))),
        is_async: false,
        captures: vec![],
    })
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
        _ => return None,
    })
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
        _ => return None,
    })
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
                    null_safe: false,
                })),
                args: vec![Argument::positional(s(from)), Argument::positional(s(to))],
                optional: false,
            })
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
        _ => None,
    }
}

/// `re.<fn>(...)` module functions → `__re_*` builtins over ecma:regexp.
fn rewrite_re_call(object: &Expression, field: &str, args: &[Argument]) -> Option<Expression> {
    if matches!(&object.kind, ExprKind::Ident(n) if n == "re") {
        let builtin = match field {
            "search" => "__re_search",
            "match" => "__re_match",
            "findall" => "__re_findall",
            "sub" => "__re_sub",
            "split" => "__re_split",
            "escape" => "__re_escape",
            "compile" => "__re_compile",
            _ => return None,
        };
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
                _ => return None,
            };
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
            null_safe: false,
        })
    };
    let member = |obj: Expression, f: &str| {
        Expression::new(ExprKind::Member {
            object: Box::new(obj),
            field: f.into(),
            null_safe: false,
        })
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
                optional: false,
            });
            // tuple(m.slice(1))
            Some(call_ident("tuple", vec![sliced]))
        }
        _ => None,
    }
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
        value: Expression::new(ExprKind::Lit(Literal::Str(v.into()))),
    };
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
        _ => return None,
    })
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
        _ => return None,
    };
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
            right: Box::new(Expression::int(mask)),
        })
    };
    let is_type = |ty: i64| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(band(0o170000)),
            right: Box::new(Expression::int(ty)),
        })
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
        _ => return None,
    })
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
        _ => return None,
    };
    Some(Literal::Str(s.into()))
}

/// Map a `string.<X>` class/function reference to its injected prelude global
/// (see [STRING_PRELUDE]). Returns `None` for constants and unknown members.
fn string_module_member(field: &str) -> Option<&'static str> {
    Some(match field {
        "Template" => "__string_Template",
        "Formatter" => "__string_Formatter",
        "capwords" => "__string_capwords",
        _ => return None,
    })
}

/// DB-API 2.0 module constants for `sqlite3` (static mount → compile-time).
fn sqlite3_module_constant(field: &str) -> Option<Literal> {
    Some(match field {
        "paramstyle" => Literal::Str("qmark".into()),
        "apilevel" => Literal::Str("2.0".into()),
        "threadsafety" => Literal::Int(1),
        "version" => Literal::Str("2.6.0".into()),
        "sqlite_version" => Literal::Str("3.40.0".into()),
        _ => return None,
    })
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
        _ => return None,
    })
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
        _ => false,
    }
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
                    optional,
                }));
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
        optional,
    }))
}

fn note_from_imported_module(name: &str) {
    PY_FROM_IMPORTED_MODULES.with(|m| m.borrow_mut().insert(name.to_string()));
}

fn is_from_imported_module(name: &str) -> bool {
    PY_FROM_IMPORTED_MODULES.with(|m| m.borrow().contains(name))
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
fn prelude_module_class(receiver: &ExprKind, field: &str) -> bool {
    let ExprKind::Ident(module) = receiver else {
        return false;
    };
    matches!(
        (module.as_str(), field),
        ("io", "StringIO")
            | ("io", "BytesIO")
            | ("configparser", "ConfigParser")
            | ("configparser", "RawConfigParser")
    )
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
    Statement::new(StmtKind::Throw {
        expr: Some(Expression::new(ExprKind::New {
            class: Box::new(Expression::new(ExprKind::Ident("ImportError".into()))),
            args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                Literal::Str(msg.into()),
            )))],
        })),
        cause: None,
    })
}

/// Python-facing names of mounted host modules that differ from the
/// canonical host export names — normalized as plain AST assignments at
/// import (`json['dumps'] = json['stringify']`), so the surface exists on
/// the runtime namespace object for reflection (dir/getattr/values) with
/// ZERO compiler/runtime machinery. JS never needs this: its names ARE
/// the canonical names.
fn py_module_renames(module: &str) -> Option<&'static [(&'static str, &'static str)]> {
    Some(match module {
        "json" => &[
            ("dumps", "stringify"),
            ("loads", "parse"),
            ("dump", "stringify"),
            ("load", "parse"),
        ],
        _ => return None,
    })
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
                null_safe: false,
            })],
            value: Expression::new(ExprKind::Index {
                object: Box::new(Expression::new(ExprKind::Ident(module.to_string()))),
                index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                    (*canonical).into(),
                )))),
                null_safe: false,
            }),
        }));
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
            )))),
        }),
        then_body: assigns,
        elifs: Vec::new(),
        else_body: None,
    })]
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
            "FunctionType",
            "LambdaType",
        ],
        "collections.abc" => &[
            "Mapping",
            "Sequence",
            "Iterable",
            "Iterator",
            "Callable",
            "Set",
            "MutableMapping",
        ],
        "json" => &["dumps", "loads", "dump", "load"],
        "zoneinfo" => &["ZoneInfo", "available_timezones", "ZoneInfoNotFoundError"],
        "glob" => &["glob", "iglob", "escape", "has_magic"],
        "fnmatch" => &["fnmatch", "fnmatchcase", "filter", "translate"],
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
        _ => return None,
    })
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
    static PY_CLASS_PARENTS: std::cell::RefCell<std::collections::HashMap<String, Vec<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
    static PY_CLASS_ATTRS: std::cell::RefCell<std::collections::HashMap<String, std::collections::HashSet<String>>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
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

fn note_class_attrs(name: &str, attrs: std::collections::HashSet<String>) {
    if !name.is_empty() {
        PY_CLASS_ATTRS.with(|m| {
            m.borrow_mut().insert(name.to_string(), attrs);
        });
    }
}

fn class_has_attr(name: &str, attr: &str) -> bool {
    PY_CLASS_ATTRS.with(|m| {
        m.borrow()
            .get(name)
            .map(|attrs| attrs.contains(attr))
            .unwrap_or(false)
    })
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
    // cross-language) via `call_or_new`. See `vybe_emitter::tuples`.
    static PY_NAMEDTUPLE_DEFS: std::cell::RefCell<std::collections::HashMap<String, NamedTupleDef>> =
        std::cell::RefCell::new(std::collections::HashMap::new());
}

#[derive(Clone)]
struct NamedTupleDef {
    type_name: String,
    fields: Vec<String>,
    /// Trailing defaults (`namedtuple(..., defaults=[...])`); apply right-aligned.
    defaults: Vec<Expression>,
}

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
            defaults: Vec::new(),
        }),
        ExprKind::Ident(name) => namedtuple_instance_def(name),
        _ => None,
    }
}

/// Positional read `recv[index]` off a namedtuple receiver.
fn namedtuple_index_read(recv: &Expression, index: usize) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(recv.clone()),
        index: Box::new(Expression::int(index as i64)),
        null_safe: false,
    })
}

/// `nt._asdict()` → an ordered dict `{field: nt[i]}`.
fn build_namedtuple_asdict(recv: &Expression, def: &NamedTupleDef) -> Expression {
    let props = def
        .fields
        .iter()
        .enumerate()
        .map(|(i, f)| ObjectProperty::KeyValue {
            key: Expression::new(ExprKind::Lit(Literal::Str(f.clone()))),
            value: namedtuple_index_read(recv, i),
        })
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
        type_name: Some(def.type_name.clone()),
    })
}

/// Extract the string value of a string-literal expression.
fn str_literal(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        _ => None,
    }
}

/// Ordered value expressions of a list/tuple literal.
fn sequence_values(expr: &Expression) -> Option<Vec<Expression>> {
    match &expr.kind {
        ExprKind::Tuple(items) | ExprKind::Set(items) => Some(items.clone()),
        ExprKind::Array(items) if items.iter().all(|e| e.key.is_none() && !e.spread) => {
            Some(items.iter().map(|e| e.value.clone()).collect())
        }
        _ => None,
    }
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
        defaults,
    })
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
            value: field_tuple,
        },
        ObjectProperty::KeyValue {
            key: Expression::string("__typename"),
            value: Expression::new(ExprKind::Lit(Literal::Str(def.type_name.clone()))),
        },
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
        type_name: Some(def.type_name.clone()),
    }
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
        _ => return None,
    })
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
        _ => return None,
    })
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
                    null_safe: false,
                })),
                Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(width)))),
            ],
            optional: false,
        })
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
            None => lit.push('%'),
        }
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
                right: Box::new(part),
            })
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
            null_safe: false,
        })
    };
    let padded = |prop: &str, width: i64| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__py_dt_pad")),
            args: vec![
                Argument::positional(field_read(prop)),
                Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(width)))),
            ],
            optional: false,
        })
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
                        by_ref: false,
                    })
                    .collect(),
            ))),
            index: Box::new(field_read("tm_wday")),
            null_safe: false,
        })
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
            _ => return None,
        };
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
                right: Box::new(part),
            })
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
        optional: false,
    })
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
                value: Expression::string("ZoneInfo"),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("key"),
                value: args[0].value.clone(),
            },
        ])),
        "available_timezones" if args.is_empty() => Some(ExprKind::Call {
            callee: Box::new(Expression::ident("set")),
            args: vec![Argument::positional(Expression::new(ExprKind::Array(
                vec![ArrayElement {
                    value: Expression::string("UTC"),
                    spread: false,
                    key: None,
                    by_ref: false,
                }],
            )))],
            optional: false,
        }),
        _ => None,
    }
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
        _ => return None,
    })
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
                None => continue,
            },
            None => i,
        };
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
            null_safe: false,
        })
    };
    let pair = Expression::new(ExprKind::Tuple(vec![pair_index(0), pair_index(1)]));
    Some(Expression::new(ExprKind::Comprehension {
        kind: ComprehensionKind::List,
        element: Box::new(pair),
        generators: vec![ComprehensionGen {
            target: Expression::new(ExprKind::Ident("__item_pair".into())),
            iter: entries,
            conditions: Vec::new(),
            is_async: false,
        }],
    }))
}

/// `dict.fromkeys(keys[, value])` → `{__k: value for __k in keys}` (value
/// defaults to `None`). Reuses the dict-comprehension lowering (which builds a
/// Map), so the result is a real dict with the right keys/order — no separate
/// classmethod machinery needed.
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
            is_async: false,
        }],
    }))
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
            optional: false,
        }))
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
                _ => return None,
            };
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
        _ => None,
    }
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
            args: vec![],
        });
        let call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(inst),
                field: "default".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                "__o".into(),
            )))],
            optional: false,
        });
        Expression::new(ExprKind::Lambda {
            params: vec![lambda_param("__o")],
            body: LambdaBody::Expr(Box::new(call)),
            is_async: false,
            captures: vec![],
        })
    } else {
        Expression::null()
    };

    let sort_keys =
        kw("sort_keys").unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(false))));
    let indent = kw("indent").unwrap_or_else(Expression::null);

    let (item_sep, kv_sep) = match kw("separators") {
        Some(sep) => match &sep.kind {
            ExprKind::Tuple(items) if items.len() == 2 => (items[0].clone(), items[1].clone()),
            _ => json_default_separators(kw("indent").is_some()),
        },
        None => json_default_separators(kw("indent").is_some()),
    };

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
        optional: false,
    }))
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
            optional: false,
        };
    }
    if let ExprKind::Ident(name) = &callee.kind {
        if let Some(def) = namedtuple_def(name) {
            return build_namedtuple_construction(&def, args);
        }
        if is_defined_class(name) {
            return ExprKind::New {
                class: Box::new(callee),
                args,
            };
        }
    }
    // Inline `namedtuple('P', 'a b')(1, 2)` — the callee is the factory call.
    if let Some(def) = parse_namedtuple_call(&callee) {
        return build_namedtuple_construction(&def, args);
    }
    ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    }
}

/// Root identifier at the base of a member/index/call chain.
fn expr_root_ident(e: &Expression) -> Option<String> {
    match &e.kind {
        ExprKind::Ident(n) => Some(n.clone()),
        ExprKind::Member { object, .. } => expr_root_ident(object),
        ExprKind::Index { object, .. } => expr_root_ident(object),
        ExprKind::Call { callee, .. } => expr_root_ident(callee),
        _ => None,
    }
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
            is_imported_module(n).then_some(module)
        }
        ExprKind::Member { object, field, .. } => {
            Some(format!("{}.{}", module_namespace_path(object)?, field))
        }
        _ => None,
    }
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
        _ => return None,
    })
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
        _ => return None,
    })
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
        _ => return None,
    })
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
        "Exception" => "Exception",
        "ValueError" => "ValueError",
        _ => return None,
    })
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
        ExprKind::Ident(name) => py_builtin_type_name(name).map(|_| "type"),
        ExprKind::Call { callee, args, .. }
            if matches!(&callee.kind, ExprKind::Ident(n) if n == "type") && args.len() == 1 =>
        {
            Some("type")
        }
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(n) if matches!(n.as_str(), "set" | "frozenset") => Some("set"),
            ExprKind::Ident(n) if n == "range" => Some("range"),
            ExprKind::Ident(n) if n == "bytes" || n == "__py_bytes_new__" => Some("bytes"),
            ExprKind::Ident(n) if n == "bytearray" => Some("bytearray"),
            ExprKind::Ident(n) if n == "complex" => Some("complex"),
            _ => None,
        },
        _ => None,
    }
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
    Some(match (sub, base) {
        (a, b) if a == b => true,
        (_, "object") => true,
        ("bool", "int") => true,
        ("ValueError", "Exception") => true,
        _ => false,
    })
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
        _ => return None,
    };
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
        _ => None,
    }
}

fn py_static_hasattr(obj: &Expression, attr: &str) -> Option<bool> {
    if let Some(type_name) = py_static_type_name(obj) {
        return Some(match (type_name, attr) {
            ("list", "append" | "extend" | "pop" | "sort" | "reverse" | "__len__") => true,
            ("tuple", "__len__") => true,
            ("dict", "keys" | "values" | "items" | "get" | "pop" | "__len__") => true,
            ("set", "add" | "discard" | "remove" | "__len__") => true,
            ("str", "upper" | "lower" | "replace" | "split" | "join" | "__len__") => true,
            ("int", "real") => true,
            _ => false,
        });
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
        _ => return None,
    })
}

/// Rewrite bare attribute reads to subscripts (see the module note above).
fn desugar_member_reads(e: Expression) -> Expression {
    match e.kind {
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => {
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
            // `sys.<const>` scalars (platform, maxsize, byteorder, …).
            if matches!(&object.kind, ExprKind::Ident(n) if n == "sys") {
                if let Some(lit) = sys_module_constant(&field) {
                    return Expression::new(ExprKind::Lit(lit));
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
                            value: Expression::new(ExprKind::Ident(name.clone())),
                        })
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
                        is_nullable: false,
                    }],
                    body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ident(
                        "__mod".into(),
                    )))),
                    is_async: false,
                    captures: vec![],
                });
            }
            // Module metadata resolves at COMPILE time — the walker knows
            // the mounts (§16.2 namespace bindings are compile-time):
            // `json.__name__` IS the import name; `__file__` is None for
            // host-backed component modules.
            if let ExprKind::Ident(module_name) = &object.kind {
                if is_imported_module(module_name) {
                    if field == "__name__" {
                        return Expression::new(ExprKind::Lit(Literal::Str(
                            module_name.clone().into(),
                        )));
                    }
                    if field == "__file__" || field == "__doc__" {
                        return Expression::new(ExprKind::Lit(Literal::Null));
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
                            optional: false,
                        });
                        let pair_index = |i: i64| {
                            Expression::new(ExprKind::Index {
                                object: Box::new(Expression::new(ExprKind::Ident(
                                    "__py_dict_pair".into(),
                                ))),
                                index: Box::new(Expression::new(ExprKind::Lit(Literal::Int(i)))),
                                null_safe: false,
                            })
                        };
                        let element = Expression::new(ExprKind::Array(vec![
                            ArrayElement {
                                key: None,
                                spread: false,
                                by_ref: false,
                                value: pair_index(0),
                            },
                            ArrayElement {
                                key: None,
                                spread: false,
                                by_ref: false,
                                value: pair_index(1),
                            },
                        ]));
                        return Expression::new(ExprKind::Comprehension {
                            kind: ComprehensionKind::Dict,
                            element: Box::new(element),
                            generators: vec![ComprehensionGen {
                                target: Expression::new(ExprKind::Ident("__py_dict_pair".into())),
                                iter: entries,
                                conditions: Vec::new(),
                                is_async: false,
                            }],
                        });
                    }
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
                        optional: false,
                    });
                }
                if let Some(names) = calendar_name_table(&full) {
                    return Expression::new(ExprKind::Array(
                        names
                            .iter()
                            .map(|n| ArrayElement {
                                value: Expression::new(ExprKind::Lit(Literal::Str((*n).into()))),
                                spread: false,
                                key: None,
                                by_ref: false,
                            })
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
                    null_safe,
                })
            } else {
                Expression::new(ExprKind::Index {
                    object: Box::new(object),
                    index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(field.into())))),
                    null_safe,
                })
            }
        }
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            // `__import__('json')` — same static mount binding as
            // importlib.import_module.
            if let ExprKind::Ident(n) = &callee.kind {
                if n == "__import__" && args.len() == 1 {
                    if let ExprKind::Lit(Literal::Str(module_name)) = &args[0].value.kind {
                        note_imported_module(module_name);
                        return Expression::new(ExprKind::Ident(module_name.clone()));
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
                            // Subscript (data) read — the stamped surface
                            // is plain properties on the namespace object.
                            return Expression::new(ExprKind::Index {
                                object: Box::new(Expression::new(ExprKind::Ident(module))),
                                index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                                    attr.to_string().into(),
                                )))),
                                null_safe: false,
                            });
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
                            _ => None,
                        },
                        _ => None,
                    };
                    if let (Some(path), ExprKind::Lit(Literal::Str(attr))) =
                        (module_path, &args[1].value.kind)
                    {
                        // Module metadata dunders always exist on a module.
                        if matches!(
                            attr.as_ref(),
                            "__name__" | "__package__" | "__doc__" | "__loader__" | "__spec__"
                        ) {
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
                if matches!(&object.kind, ExprKind::Ident(n) if n == "pkgutil")
                    && field == "iter_modules"
                {
                    return Expression::new(ExprKind::Array(vec![]));
                }
                // `types.ModuleType('name')` — a module object with its
                // `__name__` metadata.
                if matches!(&object.kind, ExprKind::Ident(n) if n == "types")
                    && field == "ModuleType"
                    && args.len() == 1
                {
                    return Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                        key: Expression::new(ExprKind::Lit(Literal::Str("__name__".into()))),
                        value: args.into_iter().next().unwrap().value,
                    }]));
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
                    if let ExprKind::Lit(Literal::Str(module_name)) = &args[0].value.kind {
                        note_imported_module(module_name);
                        return Expression::new(ExprKind::Ident(module_name.clone()));
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
                            if let ExprKind::Lit(Literal::Str(module_name)) = &args[0].value.kind {
                                return Expression::new(ExprKind::Object(vec![
                                    ObjectProperty::KeyValue {
                                        key: Expression::new(ExprKind::Lit(Literal::Str(
                                            "name".into(),
                                        ))),
                                        value: Expression::new(ExprKind::Lit(Literal::Str(
                                            module_name.clone(),
                                        ))),
                                    },
                                ]));
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
                                    null_safe: false,
                                })),
                                args: vec![],
                                optional: false,
                            })
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
                if let Some(rewritten) =
                    rewrite_sqlite_call(object, field, args.clone(), optional)
                {
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
            }
            // `string.Template(...)` / `string.Formatter()` / `string.capwords(...)`
            // — call the injected prelude global (see [STRING_PRELUDE]). Kept as a
            // real Call so keyword args (e.g. `capwords(s, sep="-")`) survive.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
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
                            optional,
                        });
                    }
                }
            }
            // Method call: keep the Member callee (method dispatch), but
            // desugar the receiver's own chain.
            let callee = match callee.kind {
                ExprKind::Member {
                    object,
                    field,
                    null_safe,
                } => Expression::new(ExprKind::Member {
                    object: Box::new(desugar_member_reads(*object)),
                    field,
                    null_safe,
                }),
                _ => desugar_member_reads(*callee),
            };
            Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args,
                optional,
            })
        }
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(desugar_member_reads(*object)),
            index,
            null_safe,
        }),
        // A module-aliased local reads AS the module (`m = json; m.dumps`),
        // and bare `__import__` is a callable value.
        ExprKind::Ident(name) => {
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
                    captures: vec![],
                });
            }
            Expression::new(ExprKind::Ident(name))
        }
        _ => e,
    }
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
                        optional: false,
                    });
                } else if let ExprKind::Member { object, field, .. } = &expr.kind {
                    // `super().__init__()` (no args) → bare `super()` parent-ctor
                    // call (see the args-carrying case below for the rationale).
                    if matches!(&object.kind, ExprKind::Super) && field == "__init__" {
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Super)),
                            args: Vec::new(),
                            optional: false,
                        });
                    } else if field == "format"
                        && let ExprKind::Lit(Literal::Str(tmpl)) = &object.kind
                        && let Some(expanded) = expand_str_format(tmpl, &[])
                    {
                        // No-arg `"literal".format()` (e.g. `'{{}}'.format()`).
                        expr = expanded;
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
                            optional: false,
                        });
                    }
                } else if matches!(&expr.kind, ExprKind::Ident(n) if n == "frozenset") {
                    // Zero-arg `frozenset()` — route to the Python builtin so the
                    // shared-compiler hack (which emits "") never fires.
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__py_frozenset")),
                        args: Vec::new(),
                        optional: false,
                    });
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
                            null_safe,
                        } = &expr.kind
                        {
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
                            if field == "sort"
                                && args.iter().any(|a| a.name.as_deref() == Some("reverse"))
                            {
                                // arr.sort(reverse=True) → arr.sort(); arr.reverse()
                                let sort_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "sort".into(),
                                        null_safe: false,
                                    })),
                                    args: vec![],
                                    optional: false,
                                });
                                let reverse_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "reverse".into(),
                                        null_safe: false,
                                    })),
                                    args: vec![],
                                    optional: false,
                                });
                                // Chain: sort then reverse. Use comma expression or sequence.
                                // Emit sort as statement, then reverse
                                expr = Expression::new(ExprKind::Sequence(vec![
                                    sort_call,
                                    reverse_call,
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
                                    is_nullable: false,
                                };
                                let filter_fn = Expression::new(ExprKind::Lambda {
                                    params: vec![param],
                                    body: LambdaBody::Expr(Box::new(Expression::new(
                                        ExprKind::Binary {
                                            op: BinOp::StrictEq,
                                            left: Box::new(Expression::new(ExprKind::Ident(
                                                "__e".into(),
                                            ))),
                                            right: Box::new(needle),
                                        },
                                    ))),
                                    is_async: false,
                                    captures: vec![],
                                });
                                let filter_call = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: object.clone(),
                                        field: "filter".into(),
                                        null_safe: false,
                                    })),
                                    args: vec![Argument::positional(filter_fn)],
                                    optional: false,
                                });
                                expr = Expression::new(ExprKind::Member {
                                    object: Box::new(filter_call),
                                    field: "length".into(),
                                    null_safe: false,
                                });
                                continue;
                            }
                            if field == "join" && args.len() == 1 {
                                let delim = object.clone();
                                let array_arg = args.into_iter().next().unwrap().value;
                                expr = Expression::new(ExprKind::Call {
                                    callee: Box::new(Expression::new(ExprKind::Member {
                                        object: Box::new(array_arg),
                                        field: "join".into(),
                                        null_safe: *null_safe,
                                    })),
                                    args: vec![Argument::positional(*delim)],
                                    optional: false,
                                });
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
                                    optional: false,
                                });
                                continue;
                            }
                        }

                        // `bytes.fromhex(s)` static constructor → Uint8Array.
                        if let ExprKind::Member { object, field, .. } = &expr.kind {
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
                                        optional: false,
                                    });
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
                                            ))),
                                        },
                                        ObjectProperty::KeyValue {
                                            key: Expression::new(ExprKind::Lit(Literal::Str(
                                                "namespace".into(),
                                            ))),
                                            value: namespace,
                                        },
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
                                        optional: false,
                                    });
                                    continue;
                                }
                                "divmod" if args.len() == 2 => {
                                    // divmod(a, b) → [a // b, a % b]
                                    let a = args[0].value.clone();
                                    let b = args[1].value.clone();
                                    expr = Expression::new(ExprKind::Array(vec![
                                        ArrayElement {
                                            key: None,
                                            spread: false,
                                            by_ref: false,
                                            value: Expression::new(ExprKind::Binary {
                                                op: BinOp::FloorDiv,
                                                left: Box::new(a.clone()),
                                                right: Box::new(b.clone()),
                                            }),
                                        },
                                        ArrayElement {
                                            key: None,
                                            spread: false,
                                            by_ref: false,
                                            value: Expression::new(ExprKind::Binary {
                                                op: BinOp::Mod,
                                                left: Box::new(a),
                                                right: Box::new(b),
                                            }),
                                        },
                                    ]));
                                    continue;
                                }
                                "callable" if args.len() == 1 => {
                                    if let Some(ok) = py_static_callable(&args[0].value) {
                                        expr = Expression::bool(ok);
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
                                                value: Expression::string(name),
                                            })
                                            .collect(),
                                        ));
                                        continue;
                                    }
                                }
                                "hasattr" if args.len() == 2 => {
                                    if let ExprKind::Lit(Literal::Str(attr)) = &args[1].value.kind {
                                        if let Some(ok) = py_static_hasattr(&args[0].value, attr) {
                                            expr = Expression::bool(ok);
                                            continue;
                                        }
                                    }
                                }
                                "int" if args.len() == 2 => {
                                    // int(s, base) → parseInt(s, base)
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "parseInt".into(),
                                        ))),
                                        args,
                                        optional: false,
                                    });
                                    continue;
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
                                                null_safe: false,
                                            })),
                                            right: Box::new(Expression::string(tag)),
                                        });
                                        continue;
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
                                                    right: Box::new(Expression::string("number")),
                                                })),
                                                right: Box::new(Expression::new(
                                                    ExprKind::Binary {
                                                        op: BinOp::StrictEq,
                                                        left: Box::new(Expression::new(
                                                            ExprKind::TypeOf(Box::new(x)),
                                                        )),
                                                        right: Box::new(Expression::string(
                                                            "boolean",
                                                        )),
                                                    },
                                                )),
                                            });
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
                                                right: Box::new(Expression::string(name)),
                                            })
                                        };
                                        let ref_test = |name: &str| {
                                            Expression::new(ExprKind::Binary {
                                                op: BinOp::InstanceOf,
                                                left: Box::new(args[0].value.clone()),
                                                right: Box::new(Expression::new(ExprKind::Ident(
                                                    name.into(),
                                                ))),
                                            })
                                        };
                                        // ref.test pushes a raw wasm i32;
                                        // materialize a real Python bool.
                                        let as_bool = |e: Expression| {
                                            Expression::new(ExprKind::Ternary {
                                                cond: Box::new(e),
                                                then: Box::new(Expression::bool(true)),
                                                else_: Box::new(Expression::bool(false)),
                                            })
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
                                                    null_safe: false,
                                                })),
                                                right: Box::new(Expression::new(ExprKind::Lit(
                                                    Literal::Null,
                                                ))),
                                            });
                                            let not_set = Expression::new(ExprKind::Unary {
                                                op: UnaryOp::Not,
                                                expr: Box::new(ref_test("Set")),
                                            });
                                            Expression::new(ExprKind::Binary {
                                                op: BinOp::And,
                                                left: Box::new(Expression::new(ExprKind::Binary {
                                                    op: BinOp::And,
                                                    left: Box::new(typeof_check("object")),
                                                    right: Box::new(not_set),
                                                })),
                                                right: Box::new(keys_probe),
                                            })
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
                                            _ => Some(as_bool(ref_test(type_name.as_str()))),
                                        };
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
                                                if let Some(ok) = py_builtin_subclass(sub, base) {
                                                    expr = Expression::bool(ok);
                                                    continue;
                                                }
                                            }
                                            ExprKind::Tuple(bases) => {
                                                let ok = bases.iter().any(|base| {
                                                    if let ExprKind::Ident(base_name) = &base.kind {
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
                                    // bool(x) → x ? True : False → ternary
                                    let x = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Ternary {
                                        cond: Box::new(x),
                                        then: Box::new(Expression::bool(true)),
                                        else_: Box::new(Expression::bool(false)),
                                    });
                                    continue;
                                }
                                "bool" if args.is_empty() => {
                                    expr = Expression::bool(false);
                                    continue;
                                }
                                "list" if args.len() == 1 => {
                                    // list(iterable) → [...iterable]
                                    let iterable = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Array(vec![ArrayElement {
                                        key: None,
                                        spread: true,
                                        by_ref: false,
                                        value: iterable,
                                    }]));
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
                                        optional: false,
                                    });
                                    continue;
                                }
                                "tuple" if args.len() == 1 => {
                                    // tuple(iterable) → [...iterable]
                                    let iterable = args[0].value.clone();
                                    expr = Expression::new(ExprKind::Array(vec![ArrayElement {
                                        key: None,
                                        spread: true,
                                        by_ref: false,
                                        value: iterable,
                                    }]));
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
                                                )),
                                            })
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
                                    let mut new_args = args;
                                    let it = new_args[0].value.clone();
                                    new_args[0].value = spread_iterable_expr(it);
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "sum".into(),
                                        ))),
                                        args: new_args,
                                        optional: false,
                                    });
                                    continue;
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
                                        optional: false,
                                    });
                                    continue;
                                }
                                "filter"
                                    if args.len() == 2
                                        && matches!(
                                            args[0].value.kind,
                                            ExprKind::Lit(Literal::Null)
                                        ) =>
                                {
                                    // filter(None, iter) → keep truthy elements
                                    // (identity predicate `lambda __e: __e`).
                                    let ident = Expression::new(ExprKind::Lambda {
                                        params: vec![Param {
                                            name: "__e".into(),
                                            type_hint: None,
                                            default: None,
                                            pass_by: PassBy::Value,
                                            is_rest: false,
                                            is_kwargs: false,
                                            is_optional: false,
                                            is_nullable: false,
                                        }],
                                        body: LambdaBody::Expr(Box::new(Expression::new(
                                            ExprKind::Ident("__e".into()),
                                        ))),
                                        is_async: false,
                                        captures: vec![],
                                    });
                                    let mut new_args = args;
                                    new_args[0] = Argument::positional(ident);
                                    expr = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "filter".into(),
                                        ))),
                                        args: new_args,
                                        optional: false,
                                    });
                                    continue;
                                }
                                "sorted" if args.len() >= 1 => {
                                    // sorted(iterable) → [...iterable].sort()
                                    // sorted(iterable, key=f) → __py_sort_by_key([...iterable], f)
                                    // sorted(..., reverse=True) → … .reverse()
                                    let iterable = args[0].value.clone();
                                    let has_reverse =
                                        args.iter().any(|a| a.name.as_deref() == Some("reverse"));
                                    let key_fn = args
                                        .iter()
                                        .find(|a| a.name.as_deref() == Some("key"))
                                        .map(|a| a.value.clone())
                                        .map(wrap_key_ident_in_lambda);
                                    let spread_array =
                                        Expression::new(ExprKind::Array(vec![ArrayElement {
                                            key: None,
                                            spread: true,
                                            by_ref: false,
                                            value: iterable,
                                        }]));
                                    let sorted = if let Some(key_fn) = key_fn {
                                        call_ident("__py_sort_by_key", vec![spread_array, key_fn])
                                    } else {
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(spread_array),
                                                field: "sort".into(),
                                                null_safe: false,
                                            })),
                                            args: vec![],
                                            optional: false,
                                        })
                                    };
                                    expr = if has_reverse {
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(sorted),
                                                field: "reverse".into(),
                                                null_safe: false,
                                            })),
                                            args: vec![],
                                            optional: false,
                                        })
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
                                        right: Box::new(n),
                                    });
                                    let scaled = Expression::new(ExprKind::Binary {
                                        op: BinOp::Mul,
                                        left: Box::new(x),
                                        right: Box::new(factor.clone()),
                                    });
                                    let rounded = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Ident(
                                            "round".into(),
                                        ))),
                                        args: vec![Argument::positional(scaled)],
                                        optional: false,
                                    });
                                    expr = Expression::new(ExprKind::Binary {
                                        op: BinOp::Div,
                                        left: Box::new(rounded),
                                        right: Box::new(factor),
                                    });
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
                                        optional: false,
                                    });
                                    expr = Expression::new(ExprKind::Binary {
                                        op: BinOp::Mod,
                                        left: Box::new(power),
                                        right: Box::new(modulus),
                                    });
                                    continue;
                                }
                                _ => {}
                            }
                        }
                        // Python `gen.throw(ExcClass)` instantiates the class so
                        // the generator's `except` matches an instance (like
                        // `raise ExcClass`). Wrap a bare uppercase-Ident arg.
                        let args = if matches!(&expr.kind, ExprKind::Member { field, .. } if field == "throw")
                            && args.first().is_some_and(|a| {
                                a.name.is_none()
                                    && matches!(&a.value.kind, ExprKind::Ident(n)
                                        if n.chars().next().is_some_and(|c| c.is_ascii_uppercase()))
                            }) {
                            let mut new_args = args;
                            let cls = new_args[0].value.clone();
                            new_args[0].value = Expression::new(ExprKind::Call {
                                callee: Box::new(cls),
                                args: vec![],
                                optional: false,
                            });
                            new_args
                        } else {
                            args
                        };
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
                        // `range(…)` → lazy generator IIFE (never materializes).
                        if matches!(&expr.kind, ExprKind::Ident(n) if n == "range") {
                            if let Some(rewritten) = lower_range_call(&args) {
                                expr = rewritten;
                                continue;
                            }
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
                        } else if prelude_module_class(&expr.kind, &field) {
                            // A prelude module's class (`io.StringIO`,
                            // `configparser.ConfigParser`, …) → the bare global
                            // class, so `mod.Class(...)` CONSTRUCTS directly. As a
                            // method call it would pass the module object as the
                            // constructor's first argument.
                            expr = Expression::new(ExprKind::Ident(field));
                        } else {
                            expr = Expression::new(ExprKind::Member {
                                object: Box::new(expr),
                                field,
                                null_safe: false,
                            });
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
                            null_safe: false,
                        });
                    }
                    _ => {
                        // Fallback: try to walk as expression
                        let val = walk_expression(children.into_iter().next().unwrap())?;
                        let val = python_index_operand(&expr, val);
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(val),
                            null_safe: false,
                        });
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
        value,
    }]))
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
            right,
        } => expr_is_python_float(left) || expr_is_python_float(right),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => expr_is_python_float(expr),
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(n) if n == "float" => true,
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
            _ => false,
        },
        _ => false,
    }
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
        optional: false,
    })
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
            right,
        } => expr_is_float_ctx(left, floats) || expr_is_float_ctx(right, floats),
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => expr_is_float_ctx(expr, floats),
        ExprKind::Call { callee, .. } if matches!(&callee.kind, ExprKind::Ident(n) if n == "__pytruediv__") => {
            true
        }
        ExprKind::Call { callee, args, .. } if matches!(&callee.kind, ExprKind::Ident(n) if is_py_arith_helper(n)) => {
            args.iter().any(|a| expr_is_float_ctx(&a.value, floats))
        }
        _ => expr_is_python_float(e),
    }
}

/// Wrap bare float-variable arguments of a `print(...)` call so they display
/// Python-float-style. (Direct float expressions were already wrapped during
/// `normalize_python_print_args`; here we catch variables tracked in `floats`.)
fn wrap_float_print_vars(e: &mut Expression, floats: &HashMap<String, bool>) {
    if let ExprKind::Call { callee, args, .. } = &mut e.kind {
        if matches!(&callee.kind, ExprKind::Ident(n) if n == "print") {
            // args[0]=sep, args[1]=end, args[2..]=items.
            for a in args.iter_mut().skip(2) {
                if a.name.is_none() && !a.spread {
                    if matches!(&a.value.kind, ExprKind::Ident(name) if *floats.get(name).unwrap_or(&false))
                    {
                        let v = std::mem::replace(&mut a.value, Expression::null());
                        a.value = wrap_float_repr(v);
                    }
                }
            }
        }
    }
}

/// Post-pass: track which local variables hold floats and wrap float-variable
/// `print` arguments. Function bodies get a fresh scope.
fn apply_float_var_repr(stmts: &mut [Statement], floats: &mut HashMap<String, bool>) {
    for stmt in stmts.iter_mut() {
        match &mut stmt.kind {
            StmtKind::Assign { targets, value } => {
                let is_f = expr_is_float_ctx(value, floats);
                if let [t] = targets.as_slice() {
                    if let ExprKind::Ident(name) = &t.kind {
                        floats.insert(name.clone(), is_f);
                    }
                }
            }
            StmtKind::Expr(e) | StmtKind::Return(Some(e)) => wrap_float_print_vars(e, floats),
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
            optional: false,
        })
    };
    // sep.join(str(x) for x in items) + end, built as a left-folded concat.
    let mut acc: Option<Expression> = None;
    for item in items {
        let piece = call_builtin1("str", item);
        acc = Some(match acc {
            None => piece,
            Some(prev) => concat(concat(prev, sep.clone()), piece),
        });
    }
    let formatted = match acc {
        Some(a) => concat(a, end),
        None => end,
    };

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(file.value.clone()),
            field: "write".into(),
            null_safe: false,
        })),
        args: vec![Argument::positional(formatted)],
        optional: false,
    }))
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
                if expr_is_python_float(&a.value) {
                    items.push(Argument::positional(wrap_float_repr(a.value)));
                } else {
                    items.push(a);
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
            by_ref: false,
        })
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
            is_async: false,
        }],
    });

    // `<list>.join(sep)` — the swapped `array.join(delim)` convention the
    // compiler expects (the source-level `delim.join(array)` swap does not run
    // on synthesized nodes).
    let joined = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(stringified),
            field: "join".into(),
            null_safe: false,
        })),
        args: vec![Argument::positional(sep)],
        optional: false,
    });

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
            spread: false,
        });
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
                    _ => val,
                };
                args.push(Argument {
                    value: val,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            } else if is_star {
                // *args — positional spread expansion.
                let val = walk_expression(ci.pop().unwrap())?;
                let val = match val.kind {
                    ExprKind::Spread(inner) => *inner,
                    _ => val,
                };
                args.push(Argument {
                    value: val,
                    name: None,
                    by_ref: false,
                    spread: true,
                });
            } else if ci.len() >= 2 && ci[0].as_rule() == Rule::identifier {
                // Check if it's keyword=value: identifier followed by expression
                // If there's an "=" between them
                let name = ci[0].as_str().to_string();
                let val = walk_expression(ci.pop().unwrap())?;
                args.push(Argument {
                    value: val,
                    name: Some(name),
                    by_ref: false,
                    spread: false,
                });
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
                        spread: true,
                    });
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
        optional: false,
    })
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
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice lower bound")?)),
        };
        let upper = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice upper bound")?)),
        };
        let step = match parts.next() {
            Some("") | None => None,
            Some(_) => Some(Box::new(exprs.next().ok_or("Missing slice step")?)),
        };
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
        _ => walk_expr_kind(inner.remove(0)),
    }
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
            generators,
        });
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
                    by_ref: false,
                })
            } else {
                Ok(ArrayElement {
                    key: None,
                    value: val,
                    spread: false,
                    by_ref: false,
                })
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
                generators,
            });
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
                    value: key,
                },
                ArrayElement {
                    key: None,
                    spread: false,
                    by_ref: false,
                    value: val,
                },
            ]));
            let generators = comp_inner
                .into_iter()
                .filter(|p| p.as_rule() == Rule::comp_clause)
                .map(walk_comp_clause)
                .collect::<Result<Vec<_>, _>>()?;
            return Ok(ExprKind::Comprehension {
                kind: ComprehensionKind::Dict,
                element: Box::new(element),
                generators,
            });
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
                        null_safe: false,
                    }),
                }));
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
                else_body: None,
            })];
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
            is_async: clause.is_async,
        })];
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
            is_sub: false,
        },
    ))));
    ExprKind::Call {
        callee: Box::new(gen_fn),
        args: Vec::new(),
        optional: false,
    }
}

/// Lower `range(stop)` / `range(start, stop)` / `range(start, stop, step)` into
/// a lazy generator IIFE so it never materializes the whole sequence — the same
/// `generators.rs` stack-switching engine the generator expression uses. Only
/// the bare positional forms are lowered; anything else keeps the profile
/// builtin.
fn lower_range_call(args: &[Argument]) -> Option<Expression> {
    if args.is_empty() || args.len() > 3 || args.iter().any(|a| a.name.is_some() || a.spread) {
        return None;
    }
    let (start, stop, step) = match args.len() {
        1 => (Expression::int(0), args[0].value.clone(), Expression::int(1)),
        2 => (
            args[0].value.clone(),
            args[1].value.clone(),
            Expression::int(1),
        ),
        _ => (
            args[0].value.clone(),
            args[1].value.clone(),
            args[2].value.clone(),
        ),
    };

    let ident = |n: &str| Expression::new(ExprKind::Ident(n.into()));
    let bin = |op, l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    let decl = |name: &str, val: Expression| {
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(name.into()),
                type_hint: None,
                init: Some(val),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        })
    };

    // while (step > 0 and i < stop) or (step < 0 and i > stop)
    let cond = bin(
        BinOp::Or,
        bin(
            BinOp::And,
            bin(BinOp::Gt, ident("__rt"), Expression::int(0)),
            bin(BinOp::Lt, ident("__ri"), ident("__re")),
        ),
        bin(
            BinOp::And,
            bin(BinOp::Lt, ident("__rt"), Expression::int(0)),
            bin(BinOp::Gt, ident("__ri"), ident("__re")),
        ),
    );
    let loop_body = vec![
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Yield(Some(
            Box::new(ident("__ri")),
        ))))),
        Statement::new(StmtKind::Assign {
            targets: vec![ident("__ri")],
            value: bin(BinOp::Add, ident("__ri"), ident("__rt")),
        }),
    ];
    let gen_body = vec![
        decl("__ri", start),
        decl("__re", stop),
        decl("__rt", step),
        Statement::new(StmtKind::While {
            cond,
            body: loop_body,
            else_body: None,
        }),
    ];
    let gen_fn = Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: Vec::new(),
            return_type: None,
            body: gen_body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: true,
            is_sub: false,
        },
    ))));
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(gen_fn),
        args: Vec::new(),
        optional: false,
    }))
}

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
            value: iter,
        }]));
    }

    Ok(ComprehensionGen {
        target,
        iter,
        conditions,
        is_async,
    })
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
                                _ => default = Some(walk_expression(c)?),
                            }
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
                                is_nullable: false,
                            });
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
        captures: Vec::new(),
    })
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
        None => (body, None),
    };
    let (name_part, conv) = match head.find('!') {
        Some(p) => (&head[..p], Some(&head[p + 1..])),
        None => (head, None),
    };

    let base = resolve_format_value(name_part, positionals, args, auto_idx)?;

    // Apply the `!r` / `!s` conversion first (Python order: convert, then spec).
    let converted = match conv {
        None => base,
        Some("r") | Some("a") => call_builtin1("repr", base),
        Some("s") => call_builtin1("str", base),
        Some(_) => return None,
    };

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
                optional: false,
            }));
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
        optional: false,
    })
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
                null_safe: false,
            });
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
                null_safe: false,
            });
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
            optional: false,
        })
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
                        null_safe: false,
                    })),
                    args: vec![],
                    optional: false,
                })
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
                optional: false,
            })
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
                null_safe: false,
            })),
            // Fill is ALWAYS explicit: rjust/ljust (str_pad_*) drop content on a
            // defaulted fill.
            args: vec![
                Argument::positional(int_lit(w)),
                Argument::positional(str_lit(&fill.to_string())),
            ],
            optional: false,
        })
    };
    match (align, width) {
        (Some(a), Some(w)) => {
            let method = match a {
                '<' => "ljust",
                '>' => "rjust",
                '^' => "center",
                _ => return None,
            };
            Some(apply(base, method, w, fill))
        }
        (Some(_), None) => Some(base), // align without width is a no-op
        (None, Some(w)) if numeric_base => {
            // Numeric base with a bare width: right-align, zero-fill if requested.
            Some(apply(base, "rjust", w, if zero { '0' } else { ' ' }))
        }
        (None, Some(_)) => None, // string base + bare width → printf (runtime default align)
        (None, None) => Some(base),
    }
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
                let Some(base) = base else { continue };
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
                    _ => base,
                };

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
                                optional: false,
                            })
                        } else {
                            call_builtin1("str", converted)
                        }
                    }
                    None if conv.is_none() => call_builtin1("str", converted),
                    None => converted,
                };
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
            _ => ArrayPatternElem::Hole,
        },
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
        _ => ArrayPatternElem::Hole,
    }
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
        end_col: ec as u32,
    }
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
            _ => return Ok(p),
        }
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
    let elements = parse_python_bytes(s)
        .into_iter()
        .map(|b| ArrayElement {
            key: None,
            spread: false,
            by_ref: false,
            value: Expression::new(ExprKind::Lit(Literal::Int(i64::from(b)))),
        })
        .collect();
    wrap_bytes(Expression::new(ExprKind::Array(elements))).kind
}

/// Build a call to a named identifier with positional args.
fn call_ident(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(name.into()))),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

/// Construct a `bytes` value from an int-array expression as a real
/// `Uint8Array` (`ObjectKind::TypedArray`). The VM handles indexing and
/// iteration natively; display is detected at runtime via `arraybuffer.isView`
/// in `emit_py_repr`, so no static bytes-tracking is needed.
fn wrap_bytes(array: Expression) -> Expression {
    call_ident("__py_bytes_new__", vec![array])
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
            _ => false,
        },
        _ => false,
    }
}

/// Decode a bytes argument (e.g. the needle of `find(b'x')`) to a latin-1
/// string; leave non-bytes args (widths, fill counts) untouched.
fn decode_bytes_arg(a: &Argument) -> Argument {
    if expr_is_python_bytes(&a.value) {
        Argument {
            value: call_ident("__vybe_bytes_decode", vec![a.value.clone()]),
            name: a.name.clone(),
            by_ref: a.by_ref,
            spread: a.spread,
        }
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
            null_safe: false,
        })),
        args: decoded_args,
        optional: false,
    });
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
    decode_python_escape_bytes(s)
        .into_iter()
        .map(char::from)
        .collect()
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
        _ => None,
    }
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
        _ => BinOp::Eq,
    }
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
        _ => BinOp::Add,
    }
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
