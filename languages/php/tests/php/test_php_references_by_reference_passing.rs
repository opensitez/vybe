use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: References & By-Reference Semantics — &$arg, function &getRef(), $a = &$b, unset($ref)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_by_reference_variable_alias() {
    let out = run_prints(
        r#"<?php
$a = 10;
$b = &$a;
$b = 20;

echo "$a - $b";
"#,
    );
    assert_eq!(out, vec!["20 - 20"]);
}

#[test]
fn test_php_function_argument_by_reference_mutation() {
    let out = run_prints(
        r#"<?php
function increment(int &$num, int $step = 1): void {
    $num += $step;
}

$val = 5;
increment($val, 10);
echo $val;
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_php_return_by_reference_from_function() {
    let out = run_prints(
        r#"<?php
$storage = 100;

function &getStorage(): int {
    global $storage;
    return $storage;
}

$ref = &getStorage();
$ref = 500;

echo $storage;
"#,
    );
    assert_eq!(out, vec!["500"]);
}

#[test]
fn test_php_unset_reference_alias_unbinding() {
    let out = run_prints(
        r#"<?php
$x = "original";
$y = &$x;
unset($y); // Unsets $y alias, $x remains
$x = "modified";

echo $x . " | " . (isset($y) ? "YES" : "NO");
"#,
    );
    assert_eq!(out, vec!["modified | NO"]);
}

#[test]
fn test_php_array_element_reference_assignment() {
    let out = run_prints(
        r#"<?php
$arr = [1, 2, 3];
$elem = &$arr[1];
$elem = 99;

echo implode(",", $arr);
"#,
    );
    assert_eq!(out, vec!["1,99,3"]);
}

#[test]
fn test_php_foreach_by_reference_dangling_reference_fix() {
    compile_ok(
        r#"<?php
$nums = [1, 2, 3];
foreach ($nums as &$v) {
    $v *= 2;
}
unset($v); // Best practice: unset reference after loop

foreach ($nums as $v) {
    echo $v . "\n";
}
"#,
    );
}

#[test]
fn test_php_class_method_return_by_reference() {
    compile_ok(
        r#"<?php
class DataContainer {
    private int $value = 42;
    public function &getValue(): int {
        return $this->value;
    }
}

$dc = new DataContainer();
$val = &$dc->getValue();
$val = 100;
echo $dc->getValue();
"#,
    );
}

#[test]
fn test_php_reference_to_property_in_object() {
    compile_ok(
        r#"<?php
class State {
    public string $status = "initial";
}

$s = new State();
$ref = &$s->status;
$ref = "updated";
echo $s->status;
"#,
    );
}

#[test]
fn test_php_swap_variables_by_reference_helper() {
    compile_ok(
        r#"<?php
function swap(&$a, &$b) {
    $tmp = $a;
    $a = $b;
    $b = $tmp;
}

$x = "first"; $y = "second";
swap($x, $y);
echo "$x $y";
"#,
    );
}

#[test]
fn test_php_array_walk_recursive_by_reference() {
    compile_ok(
        r#"<?php
$nested = ["a" => 1, "b" => [2, 3]];
array_walk_recursive($nested, function(&$val) {
    $val += 10;
});
echo $nested["a"] . " " . $nested["b"][0];
"#,
    );
}
