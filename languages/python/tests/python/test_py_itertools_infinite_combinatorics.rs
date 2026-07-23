use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Itertools Infinite & Combinatorics — count, cycle, repeat, product, permutations, combinations, groupby, islice
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_itertools_count_infinite_sequence() {
    let src = r#"
from itertools import count, islice

counter = count(start=10, step=2)
first_5 = list(islice(counter, 5))
print(first_5)
"#;
    assert_eq!(run_python(src), vec!["[10, 12, 14, 16, 18]"]);
}

#[test]
fn test_py_itertools_cycle_infinite_repeater() {
    let src = r#"
from itertools import cycle, islice

cycler = cycle(["A", "B", "C"])
sample = list(islice(cycler, 7))
print(sample)
"#;
    assert_eq!(run_python(src), vec!["['A', 'B', 'C', 'A', 'B', 'C', 'A']"]);
}

#[test]
fn test_py_itertools_repeat_fixed_times() {
    let src = r#"
from itertools import repeat

rep = list(repeat("item", 4))
print(rep)
"#;
    assert_eq!(run_python(src), vec!["['item', 'item', 'item', 'item']"]);
}

#[test]
fn test_py_itertools_product_cartesian_pairs() {
    let src = r#"
from itertools import product

p1 = list(product("AB", [1, 2]))
print(p1)

p2 = list(product([0, 1], repeat=2))
print(p2)
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
from itertools import permutations, combinations, combinations_with_replacement

items = [1, 2, 3]
print(list(permutations(items, 2)))
print(list(combinations(items, 2)))
print(list(combinations_with_replacement("AB", 2)))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[(1, 2), (1, 3), (2, 1), (2, 3), (3, 1), (3, 2)]",
            "[(1, 2), (1, 3), (2, 3)]",
            "[('A', 'A'), ('A', 'B'), ('B', 'B')]"
        ]
    );
}

#[test]
fn test_py_itertools_groupby_consecutive_runs() {
    let src = r#"
from itertools import groupby

data = [1, 1, 2, 3, 3, 3, 2, 2, 1]
grouped = [(k, list(g)) for k, g in groupby(data)]
print(grouped)
"#;
    assert_eq!(
        run_python(src),
        vec!["[(1, [1, 1]), (2, [2]), (3, [3, 3, 3]), (2, [2, 2]), (1, [1])]"]
    );
}

#[test]
fn test_py_itertools_chain_and_chain_from_iterable() {
    let src = r#"
from itertools import chain

c1 = list(chain([1, 2], [3, 4], [5]))
print(c1)

c2 = list(chain.from_iterable([[1, 2], [3, 4]]))
print(c2)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4, 5]", "[1, 2, 3, 4]"]);
}

#[test]
fn test_py_itertools_accumulate_running_totals() {
    let src = r#"
from itertools import accumulate
import operator

data = [1, 2, 3, 4, 5]
sums = list(accumulate(data))
prods = list(accumulate(data, operator.mul))

print(sums)
print(prods)
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 3, 6, 10, 15]", "[1, 2, 6, 24, 120]"]
    );
}

#[test]
fn test_py_itertools_takewhile_dropwhile_predicates() {
    let src = r#"
from itertools import takewhile, dropwhile

nums = [1, 3, 5, 2, 4, 6, 1, 3]
print(list(takewhile(lambda x: x < 5, nums)))
print(list(dropwhile(lambda x: x < 5, nums)))
"#;
    assert_eq!(run_python(src), vec!["[1, 3]", "[5, 2, 4, 6, 1, 3]"]);
}

#[test]
fn test_py_itertools_zip_longest_fillvalue() {
    let src = r#"
from itertools import zip_longest

a = [1, 2]
b = ["x", "y", "z"]

zipped = list(zip_longest(a, b, fillvalue=None))
print(zipped)
"#;
    assert_eq!(run_python(src), vec!["[(1, 'x'), (2, 'y'), (None, 'z')]"]);
}
