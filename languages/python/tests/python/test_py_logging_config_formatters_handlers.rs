use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Logging Configuration, Formatters & Handlers — getLogger, basicConfig, StreamHandler, FileHandler, Formatter, Filter, dictConfig
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_logging_basic_logger_and_levels() {
    let src = r#"
import logging, io

stream = io.StringIO()
logger = logging.getLogger("test_logger")
logger.setLevel(logging.INFO)

handler = logging.StreamHandler(stream)
handler.setFormatter(logging.Formatter("%(levelname)s:%(name)s:%(message)s"))
logger.addHandler(handler)

logger.debug("Debug msg ignored")
logger.info("Info msg logged")
logger.warning("Warning msg logged")

print(stream.getvalue().strip().splitlines())
"#;
    assert_eq!(
        run_python(src),
        vec!["['INFO:test_logger:Info msg logged', 'WARNING:test_logger:Warning msg logged']"]
    );
}

#[test]
fn test_py_logging_custom_filter_class() {
    let src = r#"
import logging, io

class AuditFilter(logging.Filter):
    def filter(self, record):
        return getattr(record, "audit", False)

stream = io.StringIO()
logger = logging.getLogger("audit_logger")
logger.setLevel(logging.DEBUG)

handler = logging.StreamHandler(stream)
handler.addFilter(AuditFilter())
logger.addHandler(handler)

logger.info("Normal message")
logger.info("Audit message", extra={"audit": True})

print("Audit message" in stream.getvalue())
print("Normal message" not in stream.getvalue())
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_logging_exception_stack_trace_formatting() {
    let src = r#"
import logging, io

stream = io.StringIO()
logger = logging.getLogger("exc_logger")
handler = logging.StreamHandler(stream)
logger.addHandler(handler)
logger.setLevel(logging.ERROR)

try:
    1 / 0
except ZeroDivisionError:
    logger.exception("Error occurred")

output = stream.getvalue()
print("Error occurred" in output)
print("ZeroDivisionError: division by zero" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_logging_dict_config_setup() {
    let src = r#"
import logging.config, io, logging

config = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "simple": {"format": "%(levelname)s - %(message)s"}
    },
    "handlers": {
        "console": {
            "class": "logging.StreamHandler",
            "level": "INFO",
            "formatter": "simple"
        }
    },
    "root": {
        "level": "INFO",
        "handlers": ["console"]
    }
}

logging.config.dictConfig(config)
root = logging.getLogger()
print(root.level == logging.INFO)
print(len(root.handlers) > 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_logging_logger_hierarchy_propagation() {
    let src = r#"
import logging, io

stream = io.StringIO()
parent = logging.getLogger("parent")
parent.setLevel(logging.INFO)
handler = logging.StreamHandler(stream)
parent.addHandler(handler)

child = logging.getLogger("parent.child")
child.info("Child message")

print("Child message" in stream.getvalue())
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_logging_file_handler_creation() {
    let src = r#"
import logging, tempfile, os

with tempfile.NamedTemporaryFile(delete=False) as f:
    fname = f.name

logger = logging.getLogger("file_logger")
logger.setLevel(logging.INFO)
fh = logging.FileHandler(fname)
logger.addHandler(fh)

logger.info("Log to file")
fh.close()

with open(fname, "r") as f:
    content = f.read()

os.unlink(fname)
print("Log to file" in content)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_logging_formatter_time_format() {
    let src = r#"
import logging

record = logging.LogRecord("test", logging.INFO, "path.py", 10, "msg", (), None)
formatter = logging.Formatter("%(asctime)s - %(message)s", datefmt="%Y-%m-%d")
formatted = formatter.format(record)
print("msg" in formatted)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_logging_get_logger_singleton() {
    let src = r#"
import logging

l1 = logging.getLogger("my_module")
l2 = logging.getLogger("my_module")
print(l1 is l2)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_logging_disable_global_level() {
    let src = r#"
import logging

logging.disable(logging.WARNING)
logger = logging.getLogger("disabled_test")
print(logger.isEnabledFor(logging.INFO) is False)
print(logger.isEnabledFor(logging.WARNING) is False)
logging.disable(logging.NOTSET)  # reset
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_logging_add_remove_handler() {
    let src = r#"
import logging

logger = logging.getLogger("handler_test")
h = logging.NullHandler()
logger.addHandler(h)
print(h in logger.handlers)
logger.removeHandler(h)
print(h in logger.handlers)
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}
