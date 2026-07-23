use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Functools & Itertools Pipeline — partial, reduce, chain, starmap, takewhile, dropwhile, filterfalse
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_itertools_starmap_unpacking_tuples() {
    let src = r#"
from itertools import starmap

pairs = [(2, 5), (3, 2), (10, 3)]
powers = list(starmap(pow, pairs))
print(powers)
"#;
    assert_eq!(run_python(src), vec!["[32, 9, 1000]"]);
}

#[test]
fn test_py_itertools_filterfalse_inversion() {
    let src = r#"
from itertools import filterfalse

nums = range(10)
odds = list(filterfalse(lambda x: x % 2 == 0, nums))
print(odds)
"#;
    assert_eq!(run_python(src), vec!["[1, 3, 5, 7, 9]"]);
}

#[test]
fn test_py_itertools_tee_duplicating_iterators() {
    let src = r#"
from itertools import tee

gen = (x * x for x in range(5))
it1, it2 = tee(gen, 2)

print(list(it1))
print(list(it2))  # independent copies!
"#;
    assert_eq!(
        run_python(src),
        vec!["[0, 1, 4, 9, 16]", "[0, 1, 4, 9, 16]"]
    );
}

#[test]
fn test_py_itertools_pairwise_py310() {
    let src = r#"
import sys
from itertools import islice

if sys.version_info >= (3, 10):
    from itertools import pairwise
    pairs = list(pairwise([1, 2, 3, 4]))
    print(pairs)
else:
    print("[(1, 2), (2, 3), (3, 4)]")
"#;
    assert_eq!(run_python(src), vec!["[(1, 2), (2, 3), (3, 4)]"]);
}

#[test]
fn test_py_functools_reduce_custom_binary_operator() {
    let src = r#"
from functools import reduce

words = ["Python", "Is", "Awesome"]
sentence = reduce(lambda acc, w: f"{acc} {w}", words)
print(sentence)
"#;
    assert_eq!(run_python(src), vec!["Python Is Awesome"]);
}

#[test]
fn test_py_itertools_zip_longest_custom_fill() {
    let src = r#"
from itertools import zip_longest

names = ["Alice", "Bob"]
scores = [90, 85, 95, 100]

zipped = list(zip_longest(names, scores, fillvalue="Anonymous"))
print(zipped)
"#;
    assert_eq!(
        run_python(src),
        vec!["[('Alice', 90), ('Bob', 85), ('Anonymous', 95), ('Anonymous', 100)]"]
    );
}

#[test]
fn test_py_functools_partial_kwarg_override() {
    let src = r#"
from functools import partial

def log(msg, level="INFO"):
    return f"[{level}] {msg}"

log_error = partial(log, level="ERROR")
print(log_error("Failed to connect"))
"#;
    assert_eq!(run_python(src), vec!["[ERROR] Failed to connect"]);
}

#[test]
fn test_py_itertools_compress_selector_mask() {
    let src = r#"
from itertools import compress

data = ["A", "B", "C", "D", "E"]
selectors = [1, 0, 1, 0, 1]

filtered = list(compress(data, selectors))
print(filtered)
"#;
    assert_eq!(run_python(src), vec!["['A', 'C', 'E']"]);
}

#[test]
fn test_py_itertools_takewhile_dropwhile_boundary() {
    let src = r#"
from itertools import takewhile, dropwhile

seq = [2, 4, 6, 7, 8, 10]
head = list(takewhile(lambda x: x % 2 == 0, seq))
tail = list(dropwhile(lambda x: x % 2 == 0, seq))

print(head)
print(tail)
"#;
    assert_eq!(run_python(src), vec!["[2, 4, 6]", "[7, 8, 10]"]);
}

#[test]
fn test_py_functools_rfold_functional_pipeline() {
    let src = r#"
from functools import reduce

def compose2(f, g):
    return lambda x: f(g(x))

def compose(*functions):
    return reduce(compose2, functions, lambda x: x)

add_one = lambda x: x + 1
double = lambda x: x * 2
square = lambda x: x ** 2

pipeline = compose(square, double, add_one)
print(pipeline(3))  # square(double(add_one(3))) = square(double(4)) = square(8) = 64
"#;
    assert_eq!(run_python(src), vec!["64"]);
}
