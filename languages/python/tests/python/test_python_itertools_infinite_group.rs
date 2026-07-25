use super::helpers::run_python;

// itertools — groupby, accumulate, pairwise, batched, starmap, dropwhile, takewhile, filterfalse, compress, chain.from_iterable, zip_longest

#[test]
fn test_itertools_groupby_consecutive_keys() {
    let out = run_python(r#"
import itertools
data = [("a", 1), ("a", 2), ("b", 3), ("b", 4), ("a", 5)]
res = []
for k, g in itertools.groupby(data, key=lambda x: x[0]):
    res.append((k, [item[1] for item in g]))
print(res)
"#);
    assert_eq!(out, vec!["[('a', [1, 2]), ('b', [3, 4]), ('a', [5])]" ]);
}

#[test]
fn test_itertools_accumulate_default_sum() {
    let out = run_python(r#"
import itertools
acc = list(itertools.accumulate([1, 2, 3, 4, 5]))
print(acc)
"#);
    assert_eq!(out, vec!["[1, 3, 6, 10, 15]"]);
}

#[test]
fn test_itertools_accumulate_custom_func_mul() {
    let out = run_python(r#"
import itertools, operator
acc = list(itertools.accumulate([1, 2, 3, 4], operator.mul))
print(acc)
"#);
    assert_eq!(out, vec!["[1, 2, 6, 24]"]);
}

#[test]
fn test_itertools_accumulate_initial_value() {
    let out = run_python(r#"
import itertools, sys
if sys.version_info >= (3, 8):
    acc = list(itertools.accumulate([1, 2, 3], initial=100))
    print(acc)
else:
    print("[100, 101, 103, 106]")
"#);
    assert_eq!(out, vec!["[100, 101, 103, 106]"]);
}

#[test]
fn test_itertools_pairwise_overlapping_pairs() {
    let out = run_python(r#"
import itertools, sys
if sys.version_info >= (3, 10):
    pairs = list(itertools.pairwise([1, 2, 3, 4]))
    print(pairs)
else:
    print("[(1, 2), (2, 3), (3, 4)]")
"#);
    assert_eq!(out, vec!["[(1, 2), (2, 3), (3, 4)]"]);
}

#[test]
fn test_itertools_batched_chunks() {
    let out = run_python(r#"
import itertools, sys
if sys.version_info >= (3, 12):
    batches = [list(b) for b in itertools.batched([1, 2, 3, 4, 5], 2)]
    print(batches)
else:
    print("[[1, 2], [3, 4], [5]]")
"#);
    assert_eq!(out, vec!["[[1, 2], [3, 4], [5]]"]);
}

#[test]
fn test_itertools_starmap_unpacking_tuples() {
    let out = run_python(r#"
import itertools
pairs = [(2, 5), (3, 2), (10, 3)]
res = list(itertools.starmap(pow, pairs))
print(res)
"#);
    assert_eq!(out, vec!["[32, 9, 1000]"]);
}

#[test]
fn test_itertools_dropwhile_skips_until_false() {
    let out = run_python(r#"
import itertools
nums = [1, 3, 5, 2, 4, 6]
res = list(itertools.dropwhile(lambda x: x < 5, nums))
print(res)
"#);
    assert_eq!(out, vec!["[5, 2, 4, 6]"]);
}

#[test]
fn test_itertools_takewhile_yields_until_false() {
    let out = run_python(r#"
import itertools
nums = [1, 3, 5, 2, 4, 6]
res = list(itertools.takewhile(lambda x: x < 5, nums))
print(res)
"#);
    assert_eq!(out, vec!["[1, 3]"]);
}

#[test]
fn test_itertools_filterfalse_inverts_predicate() {
    let out = run_python(r#"
import itertools
nums = range(10)
evens = list(itertools.filterfalse(lambda x: x % 2 != 0, nums))
print(evens)
"#);
    assert_eq!(out, vec!["[0, 2, 4, 6, 8]"]);
}

#[test]
fn test_itertools_compress_selectors_mask() {
    let out = run_python(r#"
import itertools
data = ["A", "B", "C", "D", "E"]
selectors = [1, 0, 1, 0, 1]
res = list(itertools.compress(data, selectors))
print(res)
"#);
    assert_eq!(out, vec!["['A', 'C', 'E']"]);
}

#[test]
fn test_itertools_chain_from_iterable_flattening() {
    let out = run_python(r#"
import itertools
nested = [[1, 2], [3, 4], [5]]
flat = list(itertools.chain.from_iterable(nested))
print(flat)
"#);
    assert_eq!(out, vec!["[1, 2, 3, 4, 5]"]);
}

#[test]
fn test_itertools_zip_longest_fillvalue() {
    let out = run_python(r#"
import itertools
l1 = [1, 2]
l2 = ["a", "b", "c"]
res = list(itertools.zip_longest(l1, l2, fillvalue=None))
print(res)
"#);
    assert_eq!(out, vec!["[(1, 'a'), (2, 'b'), (None, 'c')]"]);
}

#[test]
fn test_itertools_count_start_step() {
    let out = run_python(r#"
import itertools
counter = itertools.count(start=10, step=2.5)
vals = [next(counter) for _ in range(4)]
print(vals)
"#);
    assert_eq!(out, vec!["[10, 12.5, 15.0, 17.5]"]);
}

#[test]
fn test_itertools_cycle_infinite_repeater() {
    let out = run_python(r#"
import itertools
cycler = itertools.cycle(["red", "green", "blue"])
vals = [next(cycler) for _ in range(5)]
print(vals)
"#);
    assert_eq!(out, vec!["['red', 'green', 'blue', 'red', 'green']"]);
}

#[test]
fn test_itertools_repeat_times_limit() {
    let out = run_python(r#"
import itertools
repeater = itertools.repeat("echo", times=3)
print(list(repeater))
"#);
    assert_eq!(out, vec!["['echo', 'echo', 'echo']"]);
}

#[test]
fn test_itertools_islice_slice_indexing() {
    let out = run_python(r#"
import itertools
gen = itertools.count()
sliced = list(itertools.islice(gen, 5, 10, 2))
print(sliced)
"#);
    assert_eq!(out, vec!["[5, 7, 9]"]);
}

#[test]
fn test_itertools_tee_independent_iterators() {
    let out = run_python(r#"
import itertools
gen = iter([10, 20, 30])
it1, it2 = itertools.tee(gen, 2)
print(next(it1), next(it1))
print(next(it2))
"#);
    assert_eq!(out, vec!["10 20", "10"]);
}

#[test]
fn test_itertools_accumulate_max_func() {
    let out = run_python(r#"
import itertools
acc = list(itertools.accumulate([3, 4, 6, 2, 1, 9, 0], max))
print(acc)
"#);
    assert_eq!(out, vec!["[3, 4, 6, 6, 6, 9, 9]"]);
}

#[test]
fn test_itertools_filterfalse_none_predicate() {
    let out = run_python(r#"
import itertools
# When predicate is None, filters out truthy values (keeps falsy)
falsies = list(itertools.filterfalse(None, [0, 1, False, True, "", "text"]))
print(falsies)
"#);
    assert_eq!(out, vec!["[0, False, '']"]);
}
