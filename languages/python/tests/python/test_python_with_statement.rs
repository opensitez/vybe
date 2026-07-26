// Python with statement — contextlib, multiple contexts, exception suppression
use super::helpers::run_python;

#[test]
fn test_with_basic_context_manager() {
    let script = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("hello")

# file should still exist (delete=False), but closed
print(os.path.exists(path))
os.unlink(path)
"#;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_with_multiple_contexts() {
    let script = r#"
import io
a = io.StringIO()
b = io.StringIO()

with a, b:
    a.write("from a")
    b.write("from b")

print(a.getvalue())
print(b.getvalue())
"#;
    assert_eq!(run_python(script), vec!["from a", "from b"]);
}

#[test]
fn test_with_exception_propagates() {
    let script = r#"
class Ctx:
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc_val, exc_tb):
        print(f"exit: {exc_type.__name__ if exc_type else None}")
        return False  # do not suppress

try:
    with Ctx():
        raise ValueError("boom")
except ValueError:
    print("caught")
"#;
    assert_eq!(run_python(script), vec!["exit: ValueError", "caught"]);
}

#[test]
fn test_with_exception_suppressed() {
    let script = r#"
class Suppressor:
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc_val, exc_tb):
        return True  # suppress exception

with Suppressor():
    raise RuntimeError("hidden")

print("continued after suppression")
"#;
    assert_eq!(run_python(script), vec!["continued after suppression"]);
}

#[test]
fn test_contextmanager_decorator() {
    let script = r#"
from contextlib import contextmanager

@contextmanager
def managed():
    print("enter")
    yield 42
    print("exit")

with managed() as val:
    print(f"val={val}")
"#;
    assert_eq!(run_python(script), vec!["enter", "val=42", "exit"]);
}

#[test]
fn test_with_as_clause() {
    let script = r#"
import io

buf = io.StringIO("test data")
with buf as f:
    content = f.read()
print(content)
"#;
    assert_eq!(run_python(script), vec!["test data"]);
}

#[test]
fn test_contextmanager_exception_handling() {
    let script = r#"
from contextlib import contextmanager

@contextmanager
def safe():
    try:
        yield
    except ValueError as e:
        print(f"caught: {e}")

with safe():
    raise ValueError("test error")

print("after")
"#;
    assert_eq!(run_python(script), vec!["caught: test error", "after"]);
}
