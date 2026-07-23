use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: logging — loggers, levels, handlers, formatters, filters, propagation, configuration
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_logging_basic_level_filtering() {
    let src = r#"
import logging, io

stream = io.StringIO()
logging.basicConfig(stream=stream, level=logging.DEBUG, format='%(levelname)s:%(message)s')
logger = logging.getLogger("test_basic")

logger.debug("debug msg")
logger.info("info msg")
logger.warning("warn msg")
logger.error("error msg")

output = stream.getvalue()
print("DEBUG:debug msg" in output)
print("INFO:info msg" in output)
print("WARNING:warn msg" in output)
print("ERROR:error msg" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_logging_logger_hierarchy() {
    let src = r#"
import logging, io

buf = io.StringIO()
handler = logging.StreamHandler(buf)
handler.setFormatter(logging.Formatter("%(name)s:%(message)s"))

parent = logging.getLogger("app")
parent.addHandler(handler)
parent.setLevel(logging.DEBUG)
parent.propagate = False

child = logging.getLogger("app.module")
child.info("child message")

print(buf.getvalue().strip())
"#;
    assert_eq!(run_python(src), vec!["app.module:child message"]);
}

#[test]
fn test_py_logging_handler_level_override() {
    let src = r#"
import logging, io

buf = io.StringIO()
handler = logging.StreamHandler(buf)
handler.setLevel(logging.WARNING)

logger = logging.getLogger("filtered_logger")
logger.setLevel(logging.DEBUG)
logger.addHandler(handler)
logger.propagate = False

logger.debug("debug message")   # filtered out by handler
logger.warning("warning message")  # passes

lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(len(lines))
print("warning message" in lines[0])
"#;
    assert_eq!(run_python(src), vec!["1", "True"]);
}

#[test]
fn test_py_logging_formatter_with_timestamp() {
    let src = r#"
import logging, io, re

buf = io.StringIO()
handler = logging.StreamHandler(buf)
formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
handler.setFormatter(formatter)

logger = logging.getLogger("ts_test")
logger.setLevel(logging.INFO)
logger.addHandler(handler)
logger.propagate = False

logger.info("timed log")
output = buf.getvalue()
print("INFO" in output)
print("timed log" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_logging_exception_info() {
    let src = r#"
import logging, io

buf = io.StringIO()
handler = logging.StreamHandler(buf)
logger = logging.getLogger("exc_log")
logger.addHandler(handler)
logger.propagate = False
logger.setLevel(logging.ERROR)

try:
    raise ValueError("test error")
except ValueError:
    logger.exception("An error occurred")

output = buf.getvalue()
print("An error occurred" in output)
print("ValueError" in output)
print("Traceback" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_logging_dict_config() {
    let src = r#"
import logging, logging.config, io

buf = io.StringIO()

config = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "simple": {"format": "%(levelname)s - %(message)s"}
    },
    "handlers": {
        "console": {
            "class": "logging.StreamHandler",
            "formatter": "simple",
            "stream": "ext://sys.stdout"
        }
    },
    "loggers": {
        "myapp": {"level": "INFO", "handlers": ["console"], "propagate": False}
    }
}

logging.config.dictConfig(config)
logger = logging.getLogger("myapp")
# Just verify it's configured correctly
print(logger.level == logging.INFO)
print(len(logger.handlers) >= 1)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_logging_null_handler() {
    let src = r#"
import logging

# Best practice for library loggers — add NullHandler
logger = logging.getLogger("mylib")
logger.addHandler(logging.NullHandler())
logger.propagate = False

# Should not raise even without configured handlers
logger.warning("library warning")
print("ok")
"#;
    assert_eq!(run_python(src), vec!["ok"]);
}

#[test]
fn test_py_logging_memory_handler() {
    let src = r#"
import logging

buf_handler = logging.handlers.MemoryHandler if hasattr(logging, 'handlers') else None

import logging.handlers, io

buf = io.StringIO()
target = logging.StreamHandler(buf)
target.setFormatter(logging.Formatter("%(message)s"))

mem_handler = logging.handlers.MemoryHandler(capacity=5, target=target)
logger = logging.getLogger("mem_log")
logger.addHandler(mem_handler)
logger.propagate = False
logger.setLevel(logging.DEBUG)

for i in range(3):
    logger.info(f"msg{i}")

mem_handler.flush()
lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(lines)
"#;
    assert_eq!(run_python(src), vec!["['msg0', 'msg1', 'msg2']"]);
}

#[test]
fn test_py_logging_filter() {
    let src = r#"
import logging, io

class PrefixFilter(logging.Filter):
    def __init__(self, prefix):
        self.prefix = prefix

    def filter(self, record):
        return record.getMessage().startswith(self.prefix)

buf = io.StringIO()
handler = logging.StreamHandler(buf)
handler.setFormatter(logging.Formatter("%(message)s"))
handler.addFilter(PrefixFilter("ALLOW:"))

logger = logging.getLogger("filtered")
logger.addHandler(handler)
logger.propagate = False
logger.setLevel(logging.DEBUG)

logger.info("ALLOW: this goes through")
logger.info("BLOCK: this is filtered")
logger.warning("ALLOW: so does this")

lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(lines)
"#;
    assert_eq!(
        run_python(src),
        vec!["['ALLOW: this goes through', 'ALLOW: so does this']"]
    );
}

#[test]
fn test_py_logging_extra_contextual_fields() {
    let src = r#"
import logging, io

buf = io.StringIO()
handler = logging.StreamHandler(buf)
handler.setFormatter(logging.Formatter("%(user)s:%(message)s"))

logger = logging.getLogger("extra_test")
logger.addHandler(handler)
logger.propagate = False
logger.setLevel(logging.INFO)

logger.info("login attempt", extra={"user": "alice"})
logger.info("logout", extra={"user": "bob"})

lines = [l.strip() for l in buf.getvalue().strip().splitlines() if l.strip()]
print(lines)
"#;
    assert_eq!(
        run_python(src),
        vec!["['alice:login attempt', 'bob:logout']"]
    );
}
