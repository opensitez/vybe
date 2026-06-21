use super::helpers::run_prints;

// ── Spread in array literals ──────────────────────────────────

#[test]
fn spread_merges_two_arrays() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2, 3];
$b = [4, 5, 6];
$merged = [...$a, ...$b];
echo implode(',', $merged);
"#
        ),
        vec!["1,2,3,4,5,6"]
    );
}

#[test]
fn spread_with_leading_element() {
    assert_eq!(
        run_prints(
            r#"<?php
$rest = [2, 3, 4];
$full = [1, ...$rest];
echo implode(',', $full);
"#
        ),
        vec!["1,2,3,4"]
    );
}

#[test]
fn spread_with_trailing_element() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = [1, 2, 3];
$full = [...$start, 4];
echo implode(',', $full);
"#
        ),
        vec!["1,2,3,4"]
    );
}

#[test]
fn spread_three_arrays() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1]; $b = [2]; $c = [3];
$all = [...$a, ...$b, ...$c];
echo implode(',', $all);
"#
        ),
        vec!["1,2,3"]
    );
}

// ── Spread with string keys (PHP 8.1) ─────────────────────────

#[test]
fn spread_string_keyed_arrays_merges_preserving_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$defaults = ['color' => 'red', 'size' => 'M'];
$overrides = ['size' => 'L', 'weight' => 'heavy'];
$result = [...$defaults, ...$overrides];
echo $result['color'] . ',' . $result['size'] . ',' . $result['weight'];
"#
        ),
        vec!["red,L,heavy"]
    );
}

#[test]
fn spread_string_keys_later_wins() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x' => 1, 'y' => 2];
$b = ['y' => 99, 'z' => 3];
$result = [...$a, ...$b];
echo $result['y'];
"#
        ),
        vec!["99"]
    );
}

#[test]
fn spread_mixed_string_and_int_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$named = ['name' => 'Alice'];
$indexed = [1, 2, 3];
$result = [...$named, ...$indexed];
echo $result['name'] . ',' . $result[0] . ',' . $result[1];
"#
        ),
        vec!["Alice,1,2"]
    );
}

// ── Spread in function calls ──────────────────────────────────

#[test]
fn spread_in_function_call() {
    assert_eq!(
        run_prints(
            r#"<?php
function add(int $a, int $b, int $c): int { return $a + $b + $c; }
$args = [1, 2, 3];
echo add(...$args);
"#
        ),
        vec!["6"]
    );
}

#[test]
fn spread_partial_positional_args() {
    assert_eq!(
        run_prints(
            r#"<?php
function greet(string $title, string $name): string { return "$title $name"; }
$extra = ['Dr.', 'Smith'];
echo greet(...$extra);
"#
        ),
        vec!["Dr. Smith"]
    );
}

#[test]
fn spread_with_named_arg_after() {
    assert_eq!(
        run_prints(
            r#"<?php
function build(string $first, string $second, string $third): string {
    return "$first-$second-$third";
}
$parts = ['a', 'b'];
echo build(...$parts, third: 'c');
"#
        ),
        vec!["a-b-c"]
    );
}

// ── Spread with generator ─────────────────────────────────────

#[test]
fn spread_from_generator() {
    assert_eq!(
        run_prints(
            r#"<?php
function three(): Generator { yield 1; yield 2; yield 3; }
$arr = [...three()];
echo implode(',', $arr);
"#
        ),
        vec!["1,2,3"]
    );
}

// ── Spread in constructor ─────────────────────────────────────

#[test]
fn spread_in_new_constructor_call() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point { public function __construct(public int $x, public int $y) {} }
$coords = [3, 4];
$p = new Point(...$coords);
echo $p->x . ',' . $p->y;
"#
        ),
        vec!["3,4"]
    );
}

// ── array_merge vs spread performance equivalent ──────────────

#[test]
fn spread_equivalent_to_array_merge_for_indexed() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1, 2];
$b = [3, 4];
$spread = [...$a, ...$b];
$merged = array_merge($a, $b);
echo implode(',', $spread) . '|' . implode(',', $merged);
"#
        ),
        vec!["1,2,3,4|1,2,3,4"]
    );
}

// ── Spread preserves values, re-indexes numerically ──────────

#[test]
fn spread_reindexes_numeric_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [5 => 'a', 10 => 'b'];
$result = [...$a, 'c'];
echo $result[0] . ',' . $result[1] . ',' . $result[2];
"#
        ),
        vec!["a,b,c"]
    );
}

// ── Spread with empty array ───────────────────────────────────

#[test]
fn spread_empty_array_contributes_nothing() {
    assert_eq!(
        run_prints(
            r#"<?php
$empty = [];
$arr = [1, ...$empty, 2];
echo implode(',', $arr);
"#
        ),
        vec!["1,2"]
    );
}

// ── Spread in array_map argument ─────────────────────────────

#[test]
fn spread_inside_array_construction_in_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$pairs = [[1, 2], [3, 4], [5, 6]];
$sums = array_map(fn($p) => array_sum([...$p]), $pairs);
echo implode(',', $sums);
"#
        ),
        vec!["3,7,11"]
    );
}

// ── Spread string keys: configuration pattern ─────────────────

#[test]
fn spread_config_merge_defaults_with_user() {
    assert_eq!(
        run_prints(
            r#"<?php
$defaults = ['debug' => false, 'timeout' => 30, 'retries' => 3];
$user = ['timeout' => 60, 'name' => 'app'];
$config = [...$defaults, ...$user];
echo $config['debug'] === false ? 'false' : 'true';
echo ',';
echo $config['timeout'] . ',' . $config['retries'] . ',' . $config['name'];
"#
        ),
        vec!["false", ",", "60,3,app"]
    );
}

// ── Spread with array_keys preserved ──────────────────────────

#[test]
fn spread_string_keys_array_keys_preserved() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['x' => 1];
$b = ['y' => 2];
$result = [...$a, ...$b];
$keys = array_keys($result);
sort($keys);
echo implode(',', $keys);
"#
        ),
        vec!["x,y"]
    );
}

// ── Spread in nested array ────────────────────────────────────

#[test]
fn spread_inside_nested_array_literal() {
    assert_eq!(
        run_prints(
            r#"<?php
$inner = [2, 3];
$nested = [[1, ...$inner, 4], [5, 6]];
echo implode(',', $nested[0]);
"#
        ),
        vec!["1,2,3,4"]
    );
}

// ── Spread with callable-based construction ───────────────────

#[test]
fn spread_with_range_function() {
    assert_eq!(
        run_prints(
            r#"<?php
$prefix = ['start'];
$numbers = [...$prefix, ...range(1, 5)];
echo implode(',', $numbers);
"#
        ),
        vec!["start,1,2,3,4,5"]
    );
}

// ── Spread with array_values re-indexing ──────────────────────

#[test]
fn spread_of_filtered_array_reindexes() {
    assert_eq!(
        run_prints(
            r#"<?php
$filtered = array_values(array_filter([1, 2, 3, 4, 5], fn($x) => $x % 2 === 0));
$result = [0, ...$filtered, 6];
echo implode(',', $result);
"#
        ),
        vec!["0,2,4,6"]
    );
}

// ── Spread in return statement ────────────────────────────────

#[test]
fn spread_in_return_builds_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function combine(array $a, array $b): array { return [...$a, ...$b]; }
echo implode(',', combine([1, 2], [3, 4]));
"#
        ),
        vec!["1,2,3,4"]
    );
}

// ── Named args with spread ────────────────────────────────────

#[test]
fn named_args_spread_from_assoc_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function createUser(string $name, int $age, string $role = 'user'): string {
    return "$name/$age/$role";
}
$args = ['age' => 25, 'name' => 'Alice'];
echo createUser(...$args);
"#
        ),
        vec!["Alice/25/user"]
    );
}
