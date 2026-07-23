use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Context Managers (`__enter__`, `__exit__`), `@contextmanager` & Exception Suppressing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_custom_context_manager_basic_lifecycle() {
    let src = r#"
class Resource:
    def __enter__(self):
        print("enter")
        return "resource_handle"

    def __exit__(self, exc_type, exc_val, exc_tb):
        print("exit")
        return False

with Resource() as res:
    print(res)
"#;
    assert_eq!(run_python(src), vec!["enter", "resource_handle", "exit"]);
}

#[test]
fn test_py_context_manager_suppresses_exception() {
    let src = r#"
class SuppressZeroDivision:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is ZeroDivisionError:
            print("Suppression handled")
            return True # Returning True suppresses exception!
        return False

with SuppressZeroDivision():
    x = 1 / 0
print("Continued execution")
"#;
    assert_eq!(
        run_python(src),
        vec!["Suppression handled", "Continued execution"]
    );
}

#[test]
fn test_py_context_manager_propagates_unhandled_exception() {
    let src = r#"
class LoggingContext:
    def __enter__(self):
        print("start")

    def __exit__(self, exc_type, exc_val, exc_tb):
        print(f"exit with {exc_type.__name__}")
        return False # Propagate exception

try:
    with LoggingContext():
        raise ValueError("Something went wrong")
except ValueError as e:
    print(f"Caught: {e}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "start",
            "exit with ValueError",
            "Caught: Something went wrong"
        ]
    );
}

#[test]
fn test_py_contextlib_contextmanager_generator_decorator() {
    let src = r#"
from contextlib import contextmanager

@contextmanager
def managed_resource(name):
    print(f"Acquire {name}")
    try:
        yield f"HANDLE:{name}"
    finally:
        print(f"Release {name}")

with managed_resource("DB") as db:
    print(db)
"#;
    assert_eq!(
        run_python(src),
        vec!["Acquire DB", "HANDLE:DB", "Release DB"]
    );
}

#[test]
fn test_py_contextlib_contextmanager_generator_exception_handling() {
    let src = r#"
from contextlib import contextmanager

@contextmanager
def safe_block():
    print("enter safe")
    try:
        yield
    except KeyError as e:
        print(f"Caught inside generator: {e}")
    finally:
        print("exit safe")

with safe_block():
    raise KeyError("missing_key")
print("After block")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "enter safe",
            "Caught inside generator: 'missing_key'",
            "exit safe",
            "After block"
        ]
    );
}

#[test]
fn test_py_multiple_context_managers_with_statement() {
    let src = r#"
class Dummy:
    def __init__(self, tag):
        self.tag = tag
    def __enter__(self):
        print(f"enter {self.tag}")
        return self.tag
    def __exit__(self, *args):
        print(f"exit {self.tag}")

with Dummy("A") as a, Dummy("B") as b:
    print(f"inside {a} {b}")
"#;
    assert_eq!(
        run_python(src),
        vec!["enter A", "enter B", "inside A B", "exit B", "exit A"]
    ); // Exits in reverse order!
}

#[test]
fn test_py_contextlib_suppress_utility() {
    let src = r#"
from contextlib import suppress

with suppress(FileNotFoundError, KeyError):
    print("before error")
    d = {}
    x = d["missing"]
print("after suppress")
"#;
    assert_eq!(run_python(src), vec!["before error", "after suppress"]);
}

#[test]
fn test_py_contextlib_redirect_stdout() {
    let src = r#"
import io
from contextlib import redirect_stdout

f = io.StringIO()
with redirect_stdout(f):
    print("Hello to buffer!")
    print("Second line")

print("Captured:", f.getvalue().strip())
"#;
    assert_eq!(
        run_python(src),
        vec!["Captured: Hello to buffer!\nSecond line"]
    );
}

#[test]
fn test_py_contextlib_exitstack_dynamic_contexts() {
    let src = r#"
from contextlib import ExitStack

class Tracker:
    def __init__(self, name):
        self.name = name
    def __enter__(self):
        print(f"Enter {self.name}")
        return self
    def __exit__(self, *args):
        print(f"Exit {self.name}")

with ExitStack() as stack:
    t1 = stack.enter_context(Tracker("First"))
    t2 = stack.enter_context(Tracker("Second"))
    print("Inside ExitStack")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Enter First",
            "Enter Second",
            "Inside ExitStack",
            "Exit Second",
            "Exit First"
        ]
    );
}

#[test]
fn test_py_exitstack_callback_and_push() {
    let src = r#"
from contextlib import ExitStack

def cleanup(item):
    print(f"Cleanup {item}")

with ExitStack() as stack:
    stack.callback(cleanup, "Item1")
    stack.callback(cleanup, "Item2")
    print("Done stack tasks")
"#;
    assert_eq!(
        run_python(src),
        vec!["Done stack tasks", "Cleanup Item2", "Cleanup Item1"]
    );
}

#[test]
fn test_py_reentrant_context_manager() {
    let src = r#"
class Reentrant:
    def __init__(self):
        self.count = 0
    def __enter__(self):
        self.count += 1
        return self.count
    def __exit__(self, *args):
        self.count -= 1

r = Reentrant()
with r as c1:
    print("c1:", c1)
    with r as c2:
        print("c2:", c2)
"#;
    assert_eq!(run_python(src), vec!["c1: 1", "c2: 2"]);
}

#[test]
fn test_py_context_manager_reusing_as_decorator() {
    let src = r#"
from contextlib import ContextDecorator

class log_func(ContextDecorator):
    def __enter__(self):
        print("Entering function")
        return self
    def __exit__(self, *exc):
        print("Exiting function")
        return False

@log_func()
def greet():
    print("Inside greet")

greet()
"#;
    assert_eq!(
        run_python(src),
        vec!["Entering function", "Inside greet", "Exiting function"]
    );
}

#[test]
fn test_py_contextlib_nullcontext_fallback() {
    let src = r#"
from contextlib import nullcontext

def get_context(enable_cm):
    class ActiveCM:
        def __enter__(self): return "active"
        def __exit__(self, *a): pass
    return ActiveCM() if enable_cm else nullcontext("default")

with get_context(False) as val:
    print(val)
with get_context(True) as val:
    print(val)
"#;
    assert_eq!(run_python(src), vec!["default", "active"]);
}

#[test]
fn test_py_contextlib_closing_adapter() {
    let src = r#"
from contextlib import closing

class CustomStream:
    def __init__(self):
        self.closed = False
    def close(self):
        self.closed = True

stream = CustomStream()
with closing(stream):
    print("Using stream")
print("Closed:", stream.closed)
"#;
    assert_eq!(run_python(src), vec!["Using stream", "Closed: True"]);
}

#[test]
fn test_py_context_manager_exception_in_enter() {
    let src = r#"
class FailingEnter:
    def __enter__(self):
        print("enter failed")
        raise RuntimeError("Enter error")
    def __exit__(self, *args):
        print("exit should NOT be called")

try:
    with FailingEnter():
        print("inside block")
except RuntimeError as e:
    print(f"Caught: {e}")
"#;
    assert_eq!(run_python(src), vec!["enter failed", "Caught: Enter error"]);
}

#[test]
fn test_py_context_manager_exception_in_exit_overrides_original() {
    let src = r#"
class FailingExit:
    def __enter__(self):
        pass
    def __exit__(self, exc_type, exc_val, exc_tb):
        raise RuntimeError("Exit error")

try:
    with FailingExit():
        raise ValueError("Original body error")
except RuntimeError as e:
    print(f"Caught exit error: {e}")
    print(f"Context cause: {e.__context__}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Caught exit error: Exit error",
            "Context cause: Original body error"
        ]
    );
}

#[test]
fn test_py_contextlib_redirect_stderr() {
    let src = r#"
import io, sys
from contextlib import redirect_stderr

err_buf = io.StringIO()
with redirect_stderr(err_buf):
    sys.stderr.write("Custom error message\n")

print("Captured stderr:", err_buf.getvalue().strip())
"#;
    assert_eq!(
        run_python(src),
        vec!["Captured stderr: Custom error message"]
    );
}

#[test]
fn test_py_contextmanager_generator_yield_value_none() {
    let src = r#"
from contextlib import contextmanager

@contextmanager
def simple_cm():
    print("start")
    yield
    print("end")

with simple_cm() as x:
    print("x is None:", x is None)
"#;
    assert_eq!(run_python(src), vec!["start", "x is None: True", "end"]);
}

#[test]
fn test_py_contextmanager_generator_must_yield_once_error() {
    let src = r#"
from contextlib import contextmanager

@contextmanager
def no_yield_cm():
    return

try:
    with no_yield_cm():
        pass
except RuntimeError as e:
    print("RuntimeError caught: generator didn't yield")
"#;
    assert_eq!(
        run_python(src),
        vec!["RuntimeError caught: generator didn't yield"]
    );
}

#[test]
fn test_py_async_context_manager_aenter_aexit() {
    let src = r#"
import asyncio

class AsyncResource:
    async def __aenter__(self):
        print("async enter")
        return "async_res"
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        print("async exit")
        return False

async def main():
    async with AsyncResource() as res:
        print(res)

asyncio.run(main())
"#;
    assert_eq!(
        run_python(src),
        vec!["async enter", "async_res", "async exit"]
    );
}
