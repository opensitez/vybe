// Python collections.abc — Sized, Iterable, Mapping, MutableMapping, Sequence
use super::helpers::run_python;

#[test]
fn test_collections_abc_sized() {
    let script = r#"
from collections.abc import Sized
print(isinstance([1, 2, 3], Sized))
print(isinstance("hello", Sized))
print(isinstance(42, Sized))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "False"]);
}

#[test]
fn test_collections_abc_iterable() {
    let script = r#"
from collections.abc import Iterable
print(isinstance([1, 2], Iterable))
print(isinstance("abc", Iterable))
print(isinstance(123, Iterable))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "False"]);
}

#[test]
fn test_collections_abc_mapping() {
    let script = r#"
from collections.abc import Mapping
print(isinstance({'a': 1}, Mapping))
print(isinstance([1], Mapping))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_collections_abc_sequence() {
    let script = r#"
from collections.abc import Sequence
print(isinstance([1, 2], Sequence))
print(isinstance((1, 2), Sequence))
print(isinstance("abc", Sequence))
print(isinstance({1, 2}, Sequence))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True", "False"]);
}

#[test]
fn test_collections_abc_callable() {
    let script = r#"
from collections.abc import Callable
def f(): pass
print(isinstance(f, Callable))
print(isinstance(42, Callable))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_collections_abc_mutablesequence() {
    let script = r#"
from collections.abc import MutableSequence
print(isinstance([1, 2], MutableSequence))
print(isinstance((1, 2), MutableSequence))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_collections_abc_iterator() {
    let script = r#"
from collections.abc import Iterator
it = iter([1, 2, 3])
print(isinstance(it, Iterator))
print(isinstance([1, 2], Iterator))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_collections_abc_generator() {
    let script = r#"
from collections.abc import Generator
def gen():
    yield 1
g = gen()
print(isinstance(g, Generator))
"#;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_collections_abc_set() {
    let script = r#"
from collections.abc import Set
print(isinstance({1, 2}, Set))
print(isinstance(frozenset([1]), Set))
print(isinstance([1], Set))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "False"]);
}
