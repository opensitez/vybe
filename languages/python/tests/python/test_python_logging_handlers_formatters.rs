use super::helpers::run_python;

// logging — Logger, StreamHandler, Formatter, Filter, LoggerAdapter, setLevel, addHandler, addFilter, getLogger, basicConfig

#[test]
fn test_logging_logger_level_filtering() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_level")
logger.setLevel(logging.WARNING)
handler = logging.StreamHandler(stream)
logger.addHandler(handler)

logger.info("info msg")
logger.warning("warning msg")
logger.error("error msg")

output = stream.getvalue()
print("info msg" not in output)
print("warning msg" in output)
print("error msg" in output)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_logging_custom_formatter() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_fmt")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
fmt = logging.Formatter("[%(levelname)s] %(name)s: %(message)s")
handler.setFormatter(fmt)
logger.addHandler(handler)

logger.info("formatted message")
print(stream.getvalue().strip())
"#,
    );
    assert_eq!(out, vec!["[INFO] test_fmt: formatted message"]);
}

#[test]
fn test_logging_custom_filter_class() {
    let out = run_python(
        r#"
import logging, io
class SecretFilter(logging.Filter):
    def filter(self, record):
        return "secret" not in record.getMessage()

stream = io.StringIO()
logger = logging.getLogger("test_filter")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
handler.addFilter(SecretFilter())
logger.addHandler(handler)

logger.info("normal log")
logger.info("this has secret data")
logger.info("another normal log")

logs = stream.getvalue().strip().split("\n")
print(len(logs))
print("secret" not in stream.getvalue())
"#,
    );
    assert_eq!(out, vec!["2", "True"]);
}

#[test]
fn test_logging_logger_adapter_extra_context() {
    let out = run_python(
        r#"
import logging, io

class ContextAdapter(logging.LoggerAdapter):
    def process(self, msg, kwargs):
        return f"[{self.extra['user']}] {msg}", kwargs

stream = io.StringIO()
logger = logging.getLogger("test_adapter")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
handler.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(handler)

adapter = ContextAdapter(logger, {"user": "Alice"})
adapter.info("user logged in")
print(stream.getvalue().strip())
"#,
    );
    assert_eq!(out, vec!["[Alice] user logged in"]);
}

#[test]
fn test_logging_exception_formatting() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_exc")
logger.setLevel(logging.ERROR)
handler = logging.StreamHandler(stream)
logger.addHandler(handler)

try:
    1 / 0
except ZeroDivisionError:
    logger.exception("calculation failed")

output = stream.getvalue()
print("calculation failed" in output)
print("ZeroDivisionError: division by zero" in output)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_logging_child_logger_hierarchy() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
parent = logging.getLogger("parent")
parent.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
parent.addHandler(handler)

child = logging.getLogger("parent.child")
child.info("child message")

print("child message" in stream.getvalue())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_logging_propagate_disabled() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
parent = logging.getLogger("par_no_prop")
parent.setLevel(logging.INFO)
parent.addHandler(logging.StreamHandler(stream))

child = logging.getLogger("par_no_prop.child")
child.propagate = False
child.info("silent to parent")

print(stream.getvalue())
"#,
    );
    assert_eq!(out, vec![""]);
}

#[test]
fn test_logging_add_and_remove_handler() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_rem_h")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
logger.addHandler(handler)
logger.info("msg 1")
logger.removeHandler(handler)
logger.info("msg 2")

print(stream.getvalue().strip())
"#,
    );
    assert_eq!(out, vec!["msg 1"]);
}

#[test]
fn test_logging_log_levels_constants() {
    let out = run_python(
        r#"
import logging
print(logging.DEBUG < logging.INFO < logging.WARNING < logging.ERROR < logging.CRITICAL)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_logging_get_level_name() {
    let out = run_python(
        r#"
import logging
print(logging.getLevelName(logging.INFO))
print(logging.getLevelName("INFO"))
"#,
    );
    assert_eq!(out, vec!["INFO", "20"]);
}

#[test]
fn test_logging_add_custom_level() {
    let out = run_python(
        r#"
import logging
VERBOSE = 15
logging.addLevelName(VERBOSE, "VERBOSE")
print(logging.getLevelName(15))
print(logging.getLevelName("VERBOSE"))
"#,
    );
    assert_eq!(out, vec!["VERBOSE", "15"]);
}

#[test]
fn test_logging_null_handler() {
    let out = run_python(
        r#"
import logging
logger = logging.getLogger("null_test")
logger.addHandler(logging.NullHandler())
logger.info("should not raise any warning or error")
print("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_logging_file_handler() {
    let out = run_python(
        r#"
import logging, tempfile, os
with tempfile.NamedTemporaryFile(delete=False) as f:
    filepath = f.name

logger = logging.getLogger("file_test")
logger.setLevel(logging.INFO)
fh = logging.FileHandler(filepath)
logger.addHandler(fh)
logger.info("file log entry")
fh.close()

with open(filepath, "r") as r:
    content = r.read()
print("file log entry" in content)
os.unlink(filepath)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_logging_log_with_args() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_args")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
handler.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(handler)

logger.info("User %s has %d items", "Bob", 5)
print(stream.getvalue().strip())
"#,
    );
    assert_eq!(out, vec!["User Bob has 5 items"]);
}

#[test]
fn test_logging_log_with_extra_dict() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_extra")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
handler.setFormatter(logging.Formatter("%(clientip)s - %(message)s"))
logger.addHandler(handler)

logger.info("Request processed", extra={"clientip": "192.168.1.1"})
print(stream.getvalue().strip())
"#,
    );
    assert_eq!(out, vec!["192.168.1.1 - Request processed"]);
}

#[test]
fn test_logging_is_enabled_for() {
    let out = run_python(
        r#"
import logging
logger = logging.getLogger("test_enabled")
logger.setLevel(logging.WARNING)
print(logger.isEnabledFor(logging.INFO))
print(logger.isEnabledFor(logging.ERROR))
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_logging_filter_by_name() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("app.db")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
handler.addFilter(logging.Filter("app.db"))
logger.addHandler(handler)

logger.info("db log")
print("db log" in stream.getvalue())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_logging_log_method_dynamic_level() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_dyn_lvl")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
logger.addHandler(handler)

logger.log(logging.INFO, "dynamic info")
logger.log(logging.ERROR, "dynamic error")
print(stream.getvalue().strip())
"#,
    );
    assert_eq!(out, vec!["[INFO] dynamic info\n[ERROR] dynamic error"]);
}

#[test]
fn test_logging_formatter_datefmt() {
    let out = run_python(
        r#"
import logging, io
stream = io.StringIO()
logger = logging.getLogger("test_datefmt")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
fmt = logging.Formatter("%(asctime)s - %(message)s", datefmt="%Y-%m-%d")
handler.setFormatter(fmt)
logger.addHandler(handler)

logger.info("date check")
log_str = stream.getvalue()
print("date check" in log_str)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_logging_shutdown_flushes() {
    let out = run_python(
        r#"
import logging
logging.shutdown()
print("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}
