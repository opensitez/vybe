use super::helpers::run_python;

#[test]
fn test_python_contextmanager() {
    let src = r#"
from contextlib import contextmanager

@contextmanager
def sample():
    print('start')
    try:
        yield 'inside'
    finally:
        print('end')

with sample() as value:
    print(value)
"#;
    assert_eq!(run_python(src), vec!["start", "inside", "end"]);
}

#[test]
fn test_python_exit_stack() {
    let src = r#"
from contextlib import ExitStack

class A:
    def __enter__(self):
        print('in')
        return self
    def __exit__(self, exc_type, exc, tb):
        print('out')

with ExitStack() as stack:
    stack.enter_context(A())
    print('body')
"#;
    assert_eq!(run_python(src), vec!["in", "body", "out"]);
}
