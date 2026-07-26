use super::helpers::run_python;

// contextlib — AsyncExitStack, aclosing, asynccontextmanager, chdir, nullcontext, ExitStack, contextmanager, suppress, redirect_stdout, redirect_stderr

#[test]
fn test_contextlib_async_exit_stack_callbacks() {
    let out = run_python(r#"
import asyncio, contextlib

async def fn():
    cleaned_up = []
    async with contextlib.AsyncExitStack() as stack:
        stack.callback(lambda: cleaned_up.append("sync_cb"))
        async def async_cb(): cleaned_up.append("async_cb")
        stack.push_async_callback(async_cb)
    print(cleaned_up)

asyncio.run(fn())
"#);
    assert_eq!(out, vec!["['async_cb', 'sync_cb']"]);
}

#[test]
fn test_contextlib_aclosing_async_generator() {
    let out = run_python(r#"
import asyncio, contextlib

closed = [False]

async def async_gen():
    try:
        yield 1
        yield 2
    finally:
        closed[0] = True

async def fn():
    async with contextlib.aclosing(async_gen()) as gen:
        v1 = await anext(gen) if hasattr(__builtins__, 'anext') else await gen.__anext__()
        print(v1)
    print(closed[0])

asyncio.run(fn())
"#);
    assert_eq!(out, vec!["1", "True"]);
}

#[test]
fn test_contextlib_asynccontextmanager_decorator() {
    let out = run_python(r#"
import asyncio, contextlib

events = []

@contextlib.asynccontextmanager
async def resource():
    events.append("enter")
    try:
        yield "resource_data"
    finally:
        events.append("exit")

async def fn():
    async with resource() as r:
        events.append(r)
    print(events)

asyncio.run(fn())
"#);
    assert_eq!(out, vec!["['enter', 'resource_data', 'exit']"]);
}

#[test]
fn test_contextlib_chdir_temporary_directory_change() {
    let out = run_python(r#"
import contextlib, os, tempfile, sys

if sys.version_info >= (3, 11):
    orig = os.getcwd()
    with tempfile.TemporaryDirectory() as tmpdir:
        with contextlib.chdir(tmpdir):
            print(os.getcwd() == os.path.realpath(tmpdir) or os.getcwd() == tmpdir)
    print(os.getcwd() == orig)
else:
    print("True\nTrue")
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_contextlib_nullcontext_dummy_manager() {
    let out = run_python(r#"
import contextlib

with contextlib.nullcontext(enter_result="default_val") as val:
    print(val)
"#);
    assert_eq!(out, vec!["default_val"]);
}

#[test]
fn test_contextlib_exit_stack_multiple_context_managers() {
    let out = run_python(r#"
import contextlib, io

buf1 = io.StringIO()
buf2 = io.StringIO()

with contextlib.ExitStack() as stack:
    f1 = stack.enter_context(contextlib.redirect_stdout(buf1))
    f2 = stack.enter_context(contextlib.redirect_stderr(buf2))
    print("to stdout")

print(buf1.getvalue().strip())
"#);
    assert_eq!(out, vec!["to stdout"]);
}

#[test]
fn test_contextlib_contextmanager_generator_flow() {
    let out = run_python(r#"
import contextlib

state = []

@contextlib.contextmanager
def managed_state():
    state.append("open")
    try:
        yield state
    finally:
        state.append("close")

with managed_state() as s:
    s.append("work")

print(state)
"#);
    assert_eq!(out, vec!["['open', 'work', 'close']"]);
}

#[test]
fn test_contextlib_suppress_caught_exceptions() {
    let out = run_python(r#"
import contextlib

with contextlib.suppress(FileNotFoundError, KeyError):
    d = {}
    x = d["missing_key"]

print("execution continues")
"#);
    assert_eq!(out, vec!["execution continues"]);
}

#[test]
fn test_contextlib_redirect_stdout_to_buffer() {
    let out = run_python(r#"
import contextlib, io

buf = io.StringIO()
with contextlib.redirect_stdout(buf):
    print("Line 1")
    print("Line 2")

print(buf.getvalue().strip().split("\n"))
"#);
    assert_eq!(out, vec!["['Line 1', 'Line 2']"]);
}

#[test]
fn test_contextlib_redirect_stderr_to_buffer() {
    let out = run_python(r#"
import contextlib, io, sys

buf = io.StringIO()
with contextlib.redirect_stderr(buf):
    sys.stderr.write("error message\n")

print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["error message"]);
}

#[test]
fn test_contextlib_closing_calls_close_on_exit() {
    let out = run_python(r#"
import contextlib

class CustomResource:
    def __init__(self): self.is_closed = False
    def close(self): self.is_closed = True

res = CustomResource()
with contextlib.closing(res):
    print(res.is_closed)

print(res.is_closed)
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_contextlib_exit_stack_pop_all() {
    let out = run_python(r#"
import contextlib

history = []

with contextlib.ExitStack() as stack:
    stack.callback(lambda: history.append("cb1"))
    new_stack = stack.pop_all()

print(history)  # cb1 NOT called yet because popped
with new_stack:
    pass
print(history)  # now called
"#);
    assert_eq!(out, vec!["[]", "['cb1']"]);
}

#[test]
fn test_contextlib_contextmanager_exception_propagation() {
    let out = run_python(r#"
import contextlib

@contextlib.contextmanager
def handle():
    try:
        yield
    except ValueError as e:
        print(f"caught in manager: {e}")

with handle():
    raise ValueError("test_err")
"#);
    assert_eq!(out, vec!["caught in manager: test_err"]);
}

#[test]
fn test_contextlib_abstract_context_manager() {
    let out = run_python(r#"
import contextlib

class DummyCM(contextlib.AbstractContextManager):
    def __enter__(self): return "ok"
    def __exit__(self, exc_type, exc_val, exc_tb): return False

with DummyCM() as val:
    print(val)
"#);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_contextlib_abstract_async_context_manager() {
    let out = run_python(r#"
import asyncio, contextlib

class DummyAsyncCM(contextlib.AbstractAsyncContextManager):
    async def __aenter__(self): return "async_ok"
    async def __aexit__(self, exc_type, exc_val, exc_tb): return False

async def fn():
    async with DummyAsyncCM() as val:
        print(val)

asyncio.run(fn())
"#);
    assert_eq!(out, vec!["async_ok"]);
}

#[test]
fn test_contextlib_exit_stack_push_custom_exit() {
    let out = run_python(r#"
import contextlib

events = []

def custom_exit(exc_type, exc_val, exc_tb):
    events.append("custom_exit")
    return True  # suppress

with contextlib.ExitStack() as stack:
    stack.push(custom_exit)
    raise RuntimeError("boom")

print(events)
"#);
    assert_eq!(out, vec!["['custom_exit']"]);
}

#[test]
fn test_contextlib_suppress_does_not_suppress_unmatched() {
    let out = run_python(r#"
import contextlib

try:
    with contextlib.suppress(KeyError):
        raise TypeError("wrong type")
except TypeError:
    print("TypeError")
"#);
    assert_eq!(out, vec!["TypeError"]);
}

#[test]
fn test_contextlib_nullcontext_async_use() {
    let out = run_python(r#"
import asyncio, contextlib

async def fn():
    async with contextlib.nullcontext("async_val") as val:
        print(val)

asyncio.run(fn())
"#);
    assert_eq!(out, vec!["async_val"]);
}

#[test]
fn test_contextlib_contextmanager_return_value() {
    let out = run_python(r#"
import contextlib

@contextlib.contextmanager
def double(n):
    yield n * 2

with double(21) as res:
    print(res)
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_contextlib_async_exit_stack_enter_async_context() {
    let out = run_python(r#"
import asyncio, contextlib

@contextlib.asynccontextmanager
async def async_res(name):
    yield f"active_{name}"

async def fn():
    async with contextlib.AsyncExitStack() as stack:
        r1 = await stack.enter_async_context(async_res("res1"))
        r2 = await stack.enter_async_context(async_res("res2"))
        print(r1, r2)

asyncio.run(fn())
"#);
    assert_eq!(out, vec!["active_res1 active_res2"]);
}
