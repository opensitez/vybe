use super::helpers::run_prints;

// ── compact() ────────────────────────────────────────────────

#[test]
fn compact_basic_variables() {
    assert_eq!(
        run_prints(
            r#"<?php
$name = 'Alice'; $age = 30;
$arr = compact('name', 'age');
echo $arr['name'] . ':' . $arr['age'];
"#
        ),
        vec!["Alice:30"]
    );
}
#[test]
fn compact_with_array_of_names() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 1; $y = 2; $z = 3;
$vars = compact('x', 'y', 'z');
echo array_sum($vars);
"#
        ),
        vec!["6"]
    );
}
#[test]
fn compact_nested_array_argument() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = 10; $b = 20;
$result = compact(['a', 'b']);
echo $result['a'] + $result['b'];
"#
        ),
        vec!["30"]
    );
}
#[test]
fn compact_skips_undefined_var() {
    assert_eq!(
        run_prints(
            r#"<?php
$defined = 'yes';
$arr = @compact('defined', 'undefined');
echo count($arr) . ':' . $arr['defined'];
"#
        ),
        vec!["1:yes"]
    );
}
#[test]
fn compact_used_in_function_context() {
    assert_eq!(
        run_prints(
            r#"<?php
function makeUser(string $name, int $age, string $role): array {
    return compact('name', 'age', 'role');
}
$u = makeUser('Bob', 25, 'admin');
echo $u['name'] . '/' . $u['role'];
"#
        ),
        vec!["Bob/admin"]
    );
}

// ── extract() ────────────────────────────────────────────────

#[test]
fn extract_imports_to_local_scope() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['foo' => 'bar', 'num' => 42];
extract($data);
echo $foo . ':' . $num;
"#
        ),
        vec!["bar:42"]
    );
}
#[test]
fn extract_returns_count_of_vars() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = ['a' => 1, 'b' => 2, 'c' => 3];
$count = extract($data);
echo $count;
"#
        ),
        vec!["3"]
    );
}
#[test]
fn extract_overwrite_default() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 'old';
extract(['x' => 'new']);
echo $x;
"#
        ),
        vec!["new"]
    );
}
#[test]
fn extract_skip_existing() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 'original';
extract(['x' => 'override'], EXTR_SKIP);
echo $x;
"#
        ),
        vec!["original"]
    );
}
#[test]
fn extract_prefix_all() {
    assert_eq!(
        run_prints(
            r#"<?php
extract(['id' => 1, 'name' => 'Joe'], EXTR_PREFIX_ALL, 'usr');
echo $usr_id . ':' . $usr_name;
"#
        ),
        vec!["1:Joe"]
    );
}
#[test]
fn extract_prefix_conflicts_only() {
    assert_eq!(
        run_prints(
            r#"<?php
$name = 'existing';
extract(['name' => 'new', 'age' => 25], EXTR_PREFIX_SAME, 'pre');
echo $name . ':' . $pre_name . ':' . $age;
"#
        ),
        vec!["existing:new:25"]
    );
}

// ── compact + extract round-trip ─────────────────────────────

#[test]
fn compact_extract_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php
$color = 'blue'; $size = 'large'; $qty = 3;
$packed = compact('color', 'size', 'qty');
unset($color, $size, $qty);
extract($packed);
echo "$color $size x$qty";
"#
        ),
        vec!["blue large x3"]
    );
}

// ── Variable variables ────────────────────────────────────────

#[test]
fn variable_variable_basic() {
    assert_eq!(
        run_prints(r#"<?php $var = 'hello'; $name = 'var'; echo $$name; "#),
        vec!["hello"]
    );
}
#[test]
fn variable_variable_set() {
    assert_eq!(
        run_prints(r#"<?php $key = 'x'; $$key = 42; echo $x; "#),
        vec!["42"]
    );
}
#[test]
fn variable_variable_in_loop() {
    assert_eq!(
        run_prints(
            r#"<?php
foreach (['a','b','c'] as $i => $name) $$name = $i * 10;
echo $a . ',' . $b . ',' . $c;
"#
        ),
        vec!["0,10,20"]
    );
}
#[test]
fn variable_variable_with_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1,2,3];
$varname = 'arr';
echo count($$varname);
"#
        ),
        vec!["3"]
    );
}

// ── Dynamic function and method calls ────────────────────────

#[test]
fn dynamic_function_call_via_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = 'strtoupper';
echo $fn('hello');
"#
        ),
        vec!["HELLO"]
    );
}
#[test]
fn dynamic_method_call_via_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
class Calc { public function double(int $n): int { return $n * 2; } }
$method = 'double';
$c = new Calc;
echo $c->$method(7);
"#
        ),
        vec!["14"]
    );
}
#[test]
fn dynamic_static_call_via_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
class MathHelper { public static function square(int $n): int { return $n * $n; } }
$m = 'square';
echo MathHelper::$m(9);
"#
        ),
        vec!["81"]
    );
}
#[test]
fn call_user_func_with_closure() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = call_user_func(fn($x) => $x ** 2, 5);
echo $result;
"#
        ),
        vec!["25"]
    );
}
#[test]
fn call_user_func_array_spread() {
    assert_eq!(
        run_prints(
            r#"<?php
echo call_user_func_array('implode', ['-', ['a','b','c']]);
"#
        ),
        vec!["a-b-c"]
    );
}
