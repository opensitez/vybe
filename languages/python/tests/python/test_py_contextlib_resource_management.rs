use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Contextlib Resource Management — @contextmanager, ExitStack, suppress, redirect_stdout, nullcontext, closing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_contextlib_contextmanager_yield_value() {
    let src = r#"
from contextlib import contextmanager

@contextmanager
def db_transaction():
    print("BEGIN TRANSACTION")
    try:
        yield "db_connection"
        print("COMMIT")
    except Exception:
        print("ROLLBACK")
        raise

with db_transaction() as conn:
    print(f"Executing query with {conn}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "BEGIN TRANSACTION",
            "Executing query with db_connection",
            "COMMIT"
        ]
    );
}

#[test]
fn test_py_contextlib_contextmanager_exception_rollback() {
    let src = r#"
from contextlib import contextmanager

@contextmanager
def db_transaction():
    print("BEGIN TRANSACTION")
    try:
        yield
        print("COMMIT")
    except ValueError:
        print("ROLLBACK")

with db_transaction():
    print("FAILING")
    raise ValueError("Query error")
"#;
    assert_eq!(
        run_python(src),
        vec!["BEGIN TRANSACTION", "FAILING", "ROLLBACK"]
    );
}

#[test]
fn test_py_contextlib_exitstack_dynamic_resource_cleanup() {
    let src = r#"
from contextlib import ExitStack, contextmanager

@contextmanager
def acquire(name):
    print(f"Acquired {name}")
    try:
        yield name
    finally:
        print(f"Released {name}")

with ExitStack() as stack:
    resources = [stack.enter_context(acquire(f"R{i}")) for i in range(3)]
    print("Working with resources")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Acquired R0",
            "Acquired R1",
            "Acquired R2",
            "Working with resources",
            "Released R2",
            "Released R1",
            "Released R0"
        ]
    );
}

#[test]
fn test_py_contextlib_suppress_ignored_exceptions() {
    let src = r#"
from contextlib import suppress

print("before")
with suppress(FileNotFoundError, KeyError):
    raise KeyError("missing")
    print("unreachable")
print("after")
"#;
    assert_eq!(run_python(src), vec!["before", "after"]);
}

#[test]
fn test_py_contextlib_redirect_stdout_capture() {
    let src = r#"
import io
from contextlib import redirect_stdout

buf = io.StringIO()
with redirect_stdout(buf):
    print("Captured message 1")
    print("Captured message 2")

print(buf.getvalue().strip().splitlines())
"#;
    assert_eq!(
        run_python(src),
        vec!["['Captured message 1', 'Captured message 2']"]
    );
}

#[test]
fn test_py_contextlib_nullcontext_optional_lock() {
    let src = r#"
from contextlib import nullcontext

def process_data(lock=None):
    ctx = lock if lock is not None else nullcontext("dummy_res")
    with ctx as res:
        return f"Processed with {res}"

print(process_data())
"#;
    assert_eq!(run_python(src), vec!["Processed with dummy_res"]);
}

#[test]
fn test_py_contextlib_closing_wrapper() {
    let src = r#"
from contextlib import closing

class Resource:
    def __init__(self):
        self.closed = False
    def close(self):
        self.closed = True

r = Resource()
with closing(r):
    print(r.closed)

print(r.closed)
"#;
    assert_eq!(run_python(src), vec!["False", "True"]);
}

#[test]
fn test_py_contextlib_asynccontextmanager_async_with() {
    let src = r#"
import asyncio
from contextlib import asynccontextmanager

@asynccontextmanager
async def async_resource():
    print("ASYNC INIT")
    try:
        yield "async_conn"
    finally:
        print("ASYNC CLEANUP")

async def main():
    async with async_resource() as conn:
        print(f"USING {conn}")

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["ASYNC INIT", "USING async_conn", "ASYNC CLEANUP"]
    );
}

#[test]
fn test_py_contextlib_exitstack_push_callback() {
    let src = r#"
from contextlib import ExitStack

cleaned = []

with ExitStack() as stack:
    stack.callback(lambda: cleaned.append("cb1"))
    stack.callback(lambda: cleaned.append("cb2"))
    print("Work done")

print(cleaned)  # LIFO order
"#;
    assert_eq!(run_python(src), vec!["Work done", "['cb2', 'cb1']"]);
}

#[test]
fn test_py_contextlib_redirect_stderr_capture() {
    let src = r#"
import io, sys
from contextlib import redirect_stderr

buf = io.StringIO()
with redirect_stderr(buf):
    print("ERROR MSG", file=sys.stderr)

print(buf.getvalue().strip())
"#;
    assert_eq!(run_python(src), vec!["ERROR MSG"]);
}
