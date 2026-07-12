use super::helpers::run_prints;

// ── Basic list() / [] destructuring ──────────────────────────

#[test]
fn list_basic_indexed_array() {
    assert_eq!(
        run_prints(r#"<?php [$a, $b, $c] = [10, 20, 30]; echo $a . ',' . $b . ',' . $c; "#),
        vec!["10,20,30"]
    );
}
#[test]
fn list_skipping_elements() {
    assert_eq!(
        run_prints(
            r#"<?php [, $second, , $fourth] = [1, 2, 3, 4]; echo $second . ',' . $fourth; "#
        ),
        vec!["2,4"]
    );
}
#[test]
fn list_swap_variables() {
    assert_eq!(
        run_prints(r#"<?php $a = 'x'; $b = 'y'; [$a, $b] = [$b, $a]; echo $a . $b; "#),
        vec!["yx"]
    );
}
#[test]
fn list_from_function_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function minmax(array $a): array { return [min($a), max($a)]; }
[$lo, $hi] = minmax([5, 2, 8, 1, 9]);
echo "$lo,$hi";
"#
        ),
        vec!["1,9"]
    );
}
#[test]
fn list_nested_destructuring() {
    assert_eq!(
        run_prints(r#"<?php [[$x, $y], $z] = [[1, 2], 3]; echo "$x,$y,$z"; "#),
        vec!["1,2,3"]
    );
}

// ── Key-based destructuring ───────────────────────────────────

#[test]
fn list_key_based_order_independent() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['name' => 'Alice', 'age' => 30];
['age' => $age, 'name' => $name] = $data;
echo "$name is $age";
"#
        ),
        vec!["Alice is 30"]
    );
}
#[test]
fn list_key_based_partial_extract() {
    assert_eq!(
        run_prints(
            r#"<?php
$point = ['x' => 3, 'y' => 4, 'z' => 5];
['x' => $x, 'z' => $z] = $point;
echo "$x,$z";
"#
        ),
        vec!["3,5"]
    );
}
#[test]
fn list_nested_key_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['user' => ['name' => 'Bob', 'role' => 'admin']];
['user' => ['name' => $name, 'role' => $role]] = $data;
echo "$name:$role";
"#
        ),
        vec!["Bob:admin"]
    );
}

// ── foreach with list() ───────────────────────────────────────

#[test]
fn foreach_list_positional() {
    assert_eq!(
        run_prints(
            r#"<?php
$rows = [[1,'a'],[2,'b'],[3,'c']];
foreach ($rows as [$id, $label]) echo $id . $label;
"#
        ),
        vec!["1a2b3c"]
    );
}
#[test]
fn foreach_list_key_destructuring() {
    assert_eq!(
        run_prints(
            r#"<?php
$people = [['name'=>'X','score'=>10],['name'=>'Y','score'=>20]];
foreach ($people as ['name'=>$n,'score'=>$s]) echo "$n=$s ";
"#
        ),
        vec!["X=10 Y=20 "]
    );
}
#[test]
fn foreach_nested_list_skip_first() {
    assert_eq!(
        run_prints(
            r#"<?php
$pairs = [[1,2,3],[4,5,6]];
foreach ($pairs as [,$b,$c]) echo $b . $c;
"#
        ),
        vec!["2356"]
    );
}

// ── list() with various types ─────────────────────────────────

#[test]
fn list_with_string_values() {
    assert_eq!(
        run_prints(
            r#"<?php
[$first, $rest] = ['hello world split'];
echo $first;
"#
        ),
        vec!["hello world split"]
    );
}
#[test]
fn list_from_explode() {
    assert_eq!(
        run_prints(
            r#"<?php
[$year, $month, $day] = explode('-', '2024-07-15');
echo "$day/$month/$year";
"#
        ),
        vec!["15/07/2024"]
    );
}
#[test]
fn list_from_preg_match() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(\d+)-(\d+)/', '42-99', $m);
[, $a, $b] = $m;
echo $a + $b;
"#
        ),
        vec!["141"]
    );
}
#[test]
fn list_with_object_array_return() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function toArray(): array { return [$this->x, $this->y]; }
    public function __construct(public int $x, public int $y) {}
}
[$x, $y] = (new Point(3, 4))->toArray();
echo hypot($x, $y);
"#
        ),
        vec!["5"]
    );
}

// ── list() edge cases ─────────────────────────────────────────

#[test]
fn list_reassignment_replaces_var() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = 99;
[$a] = [1, 2, 3];
echo $a;
"#
        ),
        vec!["1"]
    );
}
#[test]
fn list_with_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [10, 20, 30];
[&$arr[0]] = [99];
echo $arr[0];
"#
        ),
        vec!["99"]
    );
}
#[test]
fn list_in_condition_not_used_but_valid() {
    assert_eq!(
        run_prints(
            r#"<?php
function data(): array { return [true, 42]; }
[$ok, $val] = data();
echo $ok ? $val : 'fail';
"#
        ),
        vec!["42"]
    );
}

// ── Combined patterns ─────────────────────────────────────────

#[test]
fn list_combined_with_null_coalescing() {
    assert_eq!(
        run_prints(
            r#"<?php
$response = ['status' => 200, 'body' => 'ok'];
['status' => $code, 'headers' => $hdrs] = $response + ['headers' => []];
echo $code . ':' . count($hdrs);
"#
        ),
        vec!["200:0"]
    );
}
#[test]
fn list_matrix_row_extraction() {
    assert_eq!(
        run_prints(
            r#"<?php
$matrix = [[1,2,3],[4,5,6],[7,8,9]];
[[$a,$b,$c], , [$g,$h,$i]] = $matrix;
echo $a + $i;
"#
        ),
        vec!["10"]
    );
}
