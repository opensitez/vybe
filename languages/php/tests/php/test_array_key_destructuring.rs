use super::helpers::run_prints;

// ── Short list syntax [] ──────────────────────────────────────

#[test]
fn short_list_basic_positional() {
    assert_eq!(
        run_prints(
            r#"<?php
[$a, $b, $c] = [1, 2, 3];
echo "$a,$b,$c";
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn short_list_skip_elements() {
    assert_eq!(
        run_prints(
            r#"<?php
[, $second, , $fourth] = [10, 20, 30, 40];
echo "$second,$fourth";
"#
        ),
        vec!["20,40"]
    );
}

#[test]
fn short_list_only_last_element() {
    assert_eq!(
        run_prints(
            r#"<?php
[,, $last] = ['a', 'b', 'c'];
echo $last;
"#
        ),
        vec!["c"]
    );
}

// ── Key-based destructuring ───────────────────────────────────

#[test]
fn key_based_destructuring_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
['name' => $name, 'age' => $age] = ['name' => 'Alice', 'age' => 30];
echo "$name,$age";
"#
        ),
        vec!["Alice,30"]
    );
}

#[test]
fn key_based_destructuring_order_independent() {
    assert_eq!(
        run_prints(
            r#"<?php
['y' => $y, 'x' => $x] = ['x' => 10, 'y' => 20];
echo "$x,$y";
"#
        ),
        vec!["10,20"]
    );
}

#[test]
fn key_based_destructuring_partial_extract() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['id' => 1, 'name' => 'Bob', 'role' => 'admin'];
['name' => $name, 'role' => $role] = $data;
echo "$name,$role";
"#
        ),
        vec!["Bob,admin"]
    );
}

// ── list() with keys ──────────────────────────────────────────

#[test]
fn list_with_string_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
list('first' => $a, 'last' => $b) = ['first' => 'John', 'last' => 'Doe'];
echo "$a $b";
"#
        ),
        vec!["John Doe"]
    );
}

// ── Nested destructuring ──────────────────────────────────────

#[test]
fn nested_short_list_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
[$a, [$b, $c]] = [1, [2, 3]];
echo "$a,$b,$c";
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn nested_key_based_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
['user' => ['name' => $name, 'age' => $age]] = ['user' => ['name' => 'Carol', 'age' => 25]];
echo "$name,$age";
"#
        ),
        vec!["Carol,25"]
    );
}

#[test]
fn mixed_positional_and_nested() {
    assert_eq!(
        run_prints(
            r#"<?php
[$first, [, $second]] = ['x', ['y', 'z']];
echo "$first,$second";
"#
        ),
        vec!["x,z"]
    );
}

// ── Swap via destructuring ────────────────────────────────────

#[test]
fn swap_two_variables_via_list() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = 'hello';
$b = 'world';
[$a, $b] = [$b, $a];
echo "$a,$b";
"#
        ),
        vec!["world,hello"]
    );
}

#[test]
fn swap_three_variables_via_list() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 1; $y = 2; $z = 3;
[$x, $y, $z] = [$z, $x, $y];
echo "$x,$y,$z";
"#
        ),
        vec!["3,1,2"]
    );
}

// ── foreach with list() ───────────────────────────────────────

#[test]
fn foreach_with_list_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
$matrix = [[1, 2], [3, 4], [5, 6]];
$sums = [];
foreach ($matrix as [$a, $b]) {
    $sums[] = $a + $b;
}
echo implode(',', $sums);
"#
        ),
        vec!["3,7,11"]
    );
}

#[test]
fn foreach_with_key_based_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
$users = [
    ['name' => 'Alice', 'score' => 90],
    ['name' => 'Bob', 'score' => 85],
];
$output = [];
foreach ($users as ['name' => $name, 'score' => $score]) {
    $output[] = "$name:$score";
}
echo implode(',', $output);
"#
        ),
        vec!["Alice:90,Bob:85"]
    );
}

#[test]
fn foreach_list_with_key_and_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [[1, 'a'], [2, 'b'], [3, 'c']];
$result = [];
foreach ($rows as $i => [$num, $letter]) {
    $result[] = "$i:$num$letter";
}
echo implode(',', $result);
"#
        ),
        vec!["0:1a,1:2b,2:3c"]
    );
}

// ── Destructuring from function return ───────────────────────

#[test]
fn destructure_function_return_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function getDimensions(): array { return [800, 600]; }
[$width, $height] = getDimensions();
echo "$width x $height";
"#
        ),
        vec!["800 x 600"]
    );
}

#[test]
fn destructure_named_keys_from_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function getUser(): array { return ['name' => 'Dave', 'age' => 40]; }
['name' => $name, 'age' => $age] = getUser();
echo "$name is $age";
"#
        ),
        vec!["Dave is 40"]
    );
}

// ── Destructuring with default values (via null coalescing) ───

#[test]
fn destructuring_missing_key_gives_null() {
    assert_eq!(
        run_prints(
            r#"<?php
['x' => $x, 'y' => $y] = ['x' => 10];
echo "$x," . var_export($y, true);
"#
        ),
        vec!["10,NULL"]
    );
}

// ── Reference in list ─────────────────────────────────────────

#[test]
fn list_with_reference_modifies_source() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1, 2, 3];
[&$arr[0]] = [99];
echo $arr[0];
"#
        ),
        vec!["99"]
    );
}

// ── Deeply nested array access ────────────────────────────────

#[test]
fn three_levels_nested_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
[[[[$val]]]] = [[[[42]]]];
echo $val;
"#
        ),
        vec!["42"]
    );
}

// ── Destructuring in match / complex expressions ──────────────

#[test]
fn destructuring_inside_foreach_match() {
    assert_eq!(
        run_prints(
            r#"<?php
$events = [['type' => 'click', 'x' => 10], ['type' => 'scroll', 'x' => 0]];
$result = [];
foreach ($events as ['type' => $type, 'x' => $x]) {
    $result[] = match($type) { 'click' => "click@$x", default => 'other' };
}
echo implode(',', $result);
"#
        ),
        vec!["click@10,other"]
    );
}

// ── Spread remaining elements ─────────────────────────────────

#[test]
fn destructuring_with_spread_rest() {
    assert_eq!(
        run_prints(
            r#"<?php
[$first, ...$rest] = [1, 2, 3, 4, 5];
echo $first . ',' . implode(',', $rest);
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn destructuring_first_and_last_with_spread() {
    assert_eq!(
        run_prints(
            r#"<?php
[$head, ...$middle] = ['a', 'b', 'c', 'd'];
echo $head . ',' . end($middle);
"#
        ),
        vec!["a,d"]
    );
}

// ── Destructuring objects converted to arrays ─────────────────

#[test]
fn destructuring_from_array_cast_of_object() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new stdClass();
$obj->x = 5;
$obj->y = 10;
$arr = (array)$obj;
echo $arr['x'] . ',' . $arr['y'];
"#
        ),
        vec!["5,10"]
    );
}

#[test]
fn list_assignment_ignores_excess_values() {
    assert_eq!(
        run_prints(
            r#"<?php
[$a, $b] = [1, 2, 3, 4];
echo "$a,$b";
"#
        ),
        vec!["1,2"]
    );
}

#[test]
fn list_assignment_with_reference_and_spread_mutation() {
    assert_eq!(
        run_prints(
            r#"<?php
$values = [1, 2, 3];
[$first, &$second, $third] = $values;
$second = 9;
echo $values[1] . "|" . $first . "|" . $third;
"#
        ),
        vec!["9|1|3"]
    );
}

#[test]
fn nested_destructuring_in_ternary_expression() {
    assert_eq!(
        run_prints(
            r#"<?php
$payload = true ? [10, 20] : [1, 2];
[$min, $max] = $payload;
echo $min > 5 ? "$min,$max" : "fallback";
"#
        ),
        vec!["10,20"]
    );
}

#[test]
fn key_destructuring_with_string_numeric_like_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ["0" => "zero", 1 => "one", "2" => "two", "name" => "n"];
["0" => $a, 1 => $b, "2" => $c] = $data;
echo "$a|$b|$c|{$data['name']}";
"#
        ),
        vec!["zero|one|two|n"]
    );
}

#[test]
fn list_with_single_element_and_spread_empty_tail() {
    assert_eq!(
        run_prints(
            r#"<?php
[$first, ...$rest] = [10];
echo $first . '|' . json_encode($rest);
"#
        ),
        vec!["10|[]"]
    );
}

#[test]
fn nested_list_with_fewer_values_keeps_defaults_via_null_coalesce() {
    assert_eq!(
        run_prints(
            r#"<?php
[$a, $b] = [1];
echo "$a|" . ($b ?? 'null');
"#
        ),
        vec!["1|null"]
    );
}

#[test]
fn list_with_duplicate_pattern_position_is_last_win() {
    assert_eq!(
        run_prints(
            r#"<?php
[$a, $a] = [1, 2];
echo $a;
"#
        ),
        vec!["2"]
    );
}

#[test]
fn foreach_list_with_reference_updates_source() {
    assert_eq!(
        run_prints(
            r#"<?php
$pairs = [[1], [2], [3]];
$sum = 0;
foreach ($pairs as &$pair) {
    $pair[0] *= 10;
    $sum += $pair[0];
}
echo $sum;
"#
        ),
        vec!["60"]
    );
}

#[test]
fn destructuring_via_nested_access() {
    assert_eq!(
        run_prints(
            r#"<?php
$shape = [[1, 2], [3, 4]];
[$first, $second] = $shape;
[$x, $y] = $first;
echo "$x,$y|" . ($second[0] + $second[1]);
"#
        ),
        vec!["1,2|7"]
    );
}

#[test]
fn destructuring_generator_to_list() {
    assert_eq!(
        run_prints(
            r#"<?php
function gen() {
    yield 10;
    yield 20;
}
[$a, $b] = iterator_to_array(gen());
echo "$a|$b";
"#,
        ),
        vec!["10|20"]
    );
}

#[test]
fn list_with_reference_target_without_ampersand_is_copy() {
    assert_eq!(
        run_prints(
            r#"<?php
$src = [1, 2];
[$a, $b] = $src;
$a = 9;
echo $src[0] . '|' . $a;
"#,
        ),
        vec!["1|9"]
    );
}

#[test]
fn list_with_reference_target_mutates_source() {
    assert_eq!(
        run_prints(
            r#"<?php
$src = [1, 2];
[$a, &$b] = $src;
$b = 9;
echo $src[1] . '|' . $b;
"#,
        ),
        vec!["9|9"]
    );
}

#[test]
fn list_with_more_targets_than_values_stays_missing_with_coalesce() {
    assert_eq!(
        run_prints(
            r#"<?php
[$first, $second, $third] = [5];
echo $first . '|' . ($second ?? 'nil') . '|' . ($third ?? 'nil');
"#,
        ),
        vec!["5|nil|nil"]
    );
}

#[test]
fn foreach_list_with_nullable_pairs_skips_invalid_values() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [[1, 2], null, [3, 4]];
$out = [];
foreach ($rows as $row) {
    if (!is_array($row)) {
        $out[] = 'skip';
        continue;
    }
    [$a, $b] = $row;
    $out[] = $a + $b;
}
echo implode('|', $out);
"#,
        ),
        vec!["3|skip|7"]
    );
}
