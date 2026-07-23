use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: itertools — chain, groupby, product, permutations, combinations, islice, zip_longest, accumulate, etc.
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_itertools_chain_sequences() {
    let src = r#"
import itertools

result = list(itertools.chain([1, 2], [3, 4], [5]))
print(result)

result2 = list(itertools.chain.from_iterable([[1, 2], [3, 4], [5]]))
print(result2)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4, 5]", "[1, 2, 3, 4, 5]"]);
}

#[test]
fn test_py_itertools_groupby_consecutive() {
    let src = r#"
import itertools

data = [("A", 1), ("A", 2), ("B", 3), ("A", 4)]
groups = {k: list(v) for k, v in itertools.groupby(data, key=lambda x: x[0])}
print(groups)
"#;
    assert_eq!(run_python(src), vec!["{'A': [('A', 4)], 'B': [('B', 3)]}"]); // groupby only groups consecutive!
}

#[test]
fn test_py_itertools_groupby_sorted_input() {
    let src = r#"
import itertools

words = ["apple", "ant", "bear", "bat", "cat"]
for key, group in itertools.groupby(sorted(words), key=lambda w: w[0]):
    print(f"{key}: {list(group)}")
"#;
    assert_eq!(
        run_python(src),
        vec!["a: ['ant', 'apple']", "b: ['bat', 'bear']", "c: ['cat']"]
    );
}

#[test]
fn test_py_itertools_product_cartesian() {
    let src = r#"
import itertools

pairs = list(itertools.product("AB", [1, 2]))
print(pairs)
print(list(itertools.product(range(2), repeat=2)))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[('A', 1), ('A', 2), ('B', 1), ('B', 2)]",
            "[(0, 0), (0, 1), (1, 0), (1, 1)]"
        ]
    );
}

#[test]
fn test_py_itertools_permutations_and_combinations() {
    let src = r#"
import itertools

perms = list(itertools.permutations([1, 2, 3], 2))
print(len(perms))
print(perms[0])

combs = list(itertools.combinations([1, 2, 3, 4], 2))
print(len(combs))
print(combs[0])
"#;
    assert_eq!(run_python(src), vec!["6", "(1, 2)", "6", "(1, 2)"]);
}

#[test]
fn test_py_itertools_combinations_with_replacement() {
    let src = r#"
import itertools

result = list(itertools.combinations_with_replacement("AB", 2))
print(result)
"#;
    assert_eq!(
        run_python(src),
        vec!["[('A', 'A'), ('A', 'B'), ('B', 'B')]"]
    );
}

#[test]
fn test_py_itertools_islice_window() {
    let src = r#"
import itertools

data = range(100)
first_five = list(itertools.islice(data, 5))
print(first_five)

middle = list(itertools.islice(data, 2, 7))
print(middle)

stepped = list(itertools.islice(data, 0, 10, 2))
print(stepped)
"#;
    assert_eq!(
        run_python(src),
        vec!["[0, 1, 2, 3, 4]", "[2, 3, 4, 5, 6]", "[0, 2, 4, 6, 8]"]
    );
}

#[test]
fn test_py_itertools_zip_longest() {
    let src = r#"
import itertools

result = list(itertools.zip_longest([1, 2, 3], ["a", "b"], fillvalue=0))
print(result)
"#;
    assert_eq!(run_python(src), vec!["[(1, 'a'), (2, 'b'), (3, 0)]"]);
}

#[test]
fn test_py_itertools_accumulate() {
    let src = r#"
import itertools, operator

running_sum = list(itertools.accumulate([1, 2, 3, 4, 5]))
print(running_sum)

running_product = list(itertools.accumulate([1, 2, 3, 4, 5], operator.mul))
print(running_product)
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 3, 6, 10, 15]", "[1, 2, 6, 24, 120]"]
    );
}

#[test]
fn test_py_itertools_takewhile_dropwhile() {
    let src = r#"
import itertools

data = [1, 2, 3, 4, 5, 1, 2]
print(list(itertools.takewhile(lambda x: x < 4, data)))
print(list(itertools.dropwhile(lambda x: x < 4, data)))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]", "[4, 5, 1, 2]"]);
}

#[test]
fn test_py_itertools_starmap() {
    let src = r#"
import itertools, operator

pairs = [(1, 2), (3, 4), (5, 6)]
result = list(itertools.starmap(operator.add, pairs))
print(result)
"#;
    assert_eq!(run_python(src), vec!["[3, 7, 11]"]);
}

#[test]
fn test_py_itertools_cycle_and_repeat() {
    let src = r#"
import itertools

cycled = list(itertools.islice(itertools.cycle("ABC"), 8))
print("".join(cycled))

repeated = list(itertools.repeat(42, 4))
print(repeated)
"#;
    assert_eq!(run_python(src), vec!["ABCABCAB", "[42, 42, 42, 42]"]);
}

#[test]
fn test_py_itertools_count_arithmetic_sequence() {
    let src = r#"
import itertools

evens = list(itertools.islice(itertools.count(0, 2), 5))
print(evens)

floats = list(itertools.islice(itertools.count(0.0, 0.5), 4))
print(floats)
"#;
    assert_eq!(
        run_python(src),
        vec!["[0, 2, 4, 6, 8]", "[0.0, 0.5, 1.0, 1.5]"]
    );
}

#[test]
fn test_py_itertools_tee_fork_iterator() {
    let src = r#"
import itertools

def gen():
    yield from range(4)

a, b = itertools.tee(gen(), 2)
print(list(a))
print(list(b))
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2, 3]", "[0, 1, 2, 3]"]);
}

#[test]
fn test_py_itertools_filterfalse_compress() {
    let src = r#"
import itertools

data = [1, 2, 3, 4, 5, 6]
odds = list(itertools.filterfalse(lambda x: x % 2 == 0, data))
print(odds)

selectors = [1, 0, 1, 1, 0, 1]
selected = list(itertools.compress(data, selectors))
print(selected)
"#;
    assert_eq!(run_python(src), vec!["[1, 3, 5]", "[1, 3, 4, 6]"]);
}

#[test]
fn test_py_itertools_pairwise_py310() {
    let src = r#"
import itertools, sys

if sys.version_info >= (3, 10):
    pairs = list(itertools.pairwise([1, 2, 3, 4]))
    print(pairs)
else:
    print([(1, 2), (2, 3), (3, 4)])
"#;
    assert_eq!(run_python(src), vec!["[(1, 2), (2, 3), (3, 4)]"]);
}

#[test]
fn test_py_itertools_sliding_window_pattern() {
    let src = r#"
import itertools, collections

def sliding_window(iterable, n):
    it = iter(iterable)
    window = collections.deque(itertools.islice(it, n), maxlen=n)
    if len(window) == n:
        yield tuple(window)
    for x in it:
        window.append(x)
        yield tuple(window)

print(list(sliding_window([1, 2, 3, 4, 5], 3)))
"#;
    assert_eq!(run_python(src), vec!["[(1, 2, 3), (2, 3, 4), (3, 4, 5)]"]);
}
