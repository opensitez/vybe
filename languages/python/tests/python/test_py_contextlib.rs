use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: contextlib — contextmanager, ExitStack, suppress, redirect_stdout, closing, nullcontext, AsyncContextManager
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_contextlib_contextmanager_decorator() {
    let src = r#"
from contextlib import contextmanager

log = []

@contextmanager
def managed(name):
    log.append(f"enter:{name}")
    try:
        yield f"resource:{name}"
    finally:
        log.append(f"exit:{name}")

with managed("DB") as res:
    log.append(f"use:{res}")

print(log)
"#;
    assert_eq!(
        run_python(src),
        vec!["['enter:DB', 'use:resource:DB', 'exit:DB']"]
    );
}

#[test]
fn test_py_contextlib_contextmanager_exception_handling() {
    let src = r#"
from contextlib import contextmanager

log = []

@contextmanager
def safe():
    log.append("enter")
    try:
        yield
    except ValueError as e:
        log.append(f"caught:{e}")
    finally:
        log.append("exit")

with safe():
    log.append("body")
    raise ValueError("oops")

print(log)
"#;
    assert_eq!(
        run_python(src),
        vec!["['enter', 'body', 'caught:oops', 'exit']"]
    );
}

#[test]
fn test_py_contextlib_suppress() {
    let src = r#"
from contextlib import suppress

log = []

with suppress(ValueError, TypeError):
    log.append("before")
    raise ValueError("ignored")
    log.append("unreachable")

log.append("after")
print(log)
"#;
    assert_eq!(run_python(src), vec!["['before', 'after']"]);
}

#[test]
fn test_py_contextlib_exit_stack() {
    let src = r#"
from contextlib import ExitStack, contextmanager

log = []

@contextmanager
def res(name):
    log.append(f"open:{name}")
    yield name
    log.append(f"close:{name}")

with ExitStack() as stack:
    for name in ["A", "B", "C"]:
        stack.enter_context(res(name))
    log.append("using")

print(log)
"#;
    assert_eq!(
        run_python(src),
        vec!["['open:A', 'open:B', 'open:C', 'using', 'close:C', 'close:B', 'close:A']"]
    );
}

#[test]
fn test_py_contextlib_redirect_stdout() {
    let src = r#"
from contextlib import redirect_stdout
import io

buf = io.StringIO()
with redirect_stdout(buf):
    print("captured output")
    print(42)

print(buf.getvalue().strip())
print("back to real stdout")
"#;
    assert_eq!(
        run_python(src),
        vec!["captured output\n42", "back to real stdout"]
    );
}

#[test]
fn test_py_contextlib_redirect_stderr() {
    let src = r#"
from contextlib import redirect_stderr
import io, sys

buf = io.StringIO()
with redirect_stderr(buf):
    print("error message", file=sys.stderr)

print(buf.getvalue().strip())
"#;
    assert_eq!(run_python(src), vec!["error message"]);
}

#[test]
fn test_py_contextlib_nullcontext() {
    let src = r#"
from contextlib import nullcontext

def process(lock=None):
    ctx = lock if lock is not None else nullcontext()
    with ctx:
        return "processed"

print(process())
print(process(nullcontext()))
"#;
    assert_eq!(run_python(src), vec!["processed", "processed"]);
}

#[test]
fn test_py_contextlib_closing() {
    let src = r#"
from contextlib import closing

class Resource:
    def __init__(self):
        self.open = True

    def close(self):
        self.open = False

r = Resource()
with closing(r):
    print(r.open)

print(r.open)  # closed after exiting
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_contextlib_asynccontextmanager() {
    let src = r#"
import asyncio
from contextlib import asynccontextmanager

log = []

@asynccontextmanager
async def async_res(name):
    log.append(f"enter:{name}")
    yield name
    log.append(f"exit:{name}")

async def main():
    async with async_res("conn") as r:
        log.append(f"use:{r}")
    print(log)

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["['enter:conn', 'use:conn', 'exit:conn']"]
    );
}

#[test]
fn test_py_contextlib_exit_stack_callback() {
    let src = r#"
from contextlib import ExitStack

log = []

with ExitStack() as stack:
    stack.callback(log.append, "callback_1")
    stack.callback(log.append, "callback_2")
    log.append("body")

print(log)  # callbacks run in LIFO order
"#;
    assert_eq!(
        run_python(src),
        vec!["['body', 'callback_2', 'callback_1']"]
    );
}

#[test]
fn test_py_contextlib_nested_context_managers() {
    let src = r#"
import tempfile, os
from contextlib import contextmanager, ExitStack

@contextmanager
def temp_file():
    import tempfile
    f = tempfile.NamedTemporaryFile(delete=False, mode='w', suffix='.txt')
    try:
        yield f
    finally:
        f.close()
        os.unlink(f.name)

with ExitStack() as stack:
    f1 = stack.enter_context(temp_file())
    f2 = stack.enter_context(temp_file())
    f1.write("file1")
    f2.write("file2")
    print(f1.name != f2.name)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
