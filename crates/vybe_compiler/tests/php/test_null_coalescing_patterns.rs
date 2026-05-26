use super::helpers::run_prints;

// ── Null coalescing ?? ────────────────────────────────────────

#[test] fn null_coalescing_null_gives_default() {
    assert_eq!(run_prints(r#"<?php $a = null; echo $a ?? 'default'; "#), vec!["default"]);
}
#[test] fn null_coalescing_false_is_not_null() {
    assert_eq!(run_prints(r#"<?php $a = false; echo ($a ?? 'x') === false ? 'kept' : 'replaced'; "#), vec!["kept"]);
}
#[test] fn null_coalescing_zero_is_not_null() {
    assert_eq!(run_prints(r#"<?php $a = 0; echo $a ?? 99; "#), vec!["0"]);
}
#[test] fn null_coalescing_chained_finds_first_non_null() {
    assert_eq!(run_prints(r#"<?php echo $a ?? $b ?? 'last'; "#), vec!["last"]);
}
#[test] fn null_coalescing_nested_array_key_missing() {
    assert_eq!(run_prints(r#"<?php $d = ['user' => ['name' => 'Al']]; echo $d['user']['email'] ?? 'none'; "#), vec!["none"]);
}

// ── Null coalescing assignment ??= ────────────────────────────

#[test] fn null_coalescing_assign_sets_when_null() {
    assert_eq!(run_prints(r#"<?php $x = null; $x ??= 42; echo $x; "#), vec!["42"]);
}
#[test] fn null_coalescing_assign_skips_when_set() {
    assert_eq!(run_prints(r#"<?php $x = 7; $x ??= 42; echo $x; "#), vec!["7"]);
}
#[test] fn null_coalescing_assign_array_missing_key() {
    assert_eq!(run_prints(r#"<?php $cfg = []; $cfg['ttl'] ??= 30; echo $cfg['ttl']; "#), vec!["30"]);
}
#[test] fn null_coalescing_assign_counter_idiom() {
    assert_eq!(run_prints(r#"<?php
$c = [];
foreach (['a','b','a'] as $v) { $c[$v] ??= 0; $c[$v]++; }
echo $c['a'] . ',' . $c['b'];
"#), vec!["2,1"]);
}

// ── Null safe operator ?-> ────────────────────────────────────

#[test] fn null_safe_short_circuits_on_null() {
    assert_eq!(run_prints(r#"<?php
class User { public ?Address $addr = null; }
class Address { public string $city = 'Paris'; }
$u = new User;
echo $u?->addr?->city ?? 'unknown';
"#), vec!["unknown"]);
}
#[test] fn null_safe_traverses_when_not_null() {
    assert_eq!(run_prints(r#"<?php
class User { public ?Address $addr; }
class Address { public string $city; }
$u = new User; $u->addr = new Address; $u->addr->city = 'Tokyo';
echo $u?->addr?->city;
"#), vec!["Tokyo"]);
}
#[test] fn null_safe_method_call() {
    assert_eq!(run_prints(r#"<?php
class Repo { public function find(): ?string { return 'item'; } }
$r = new Repo;
echo $r?->find();
"#), vec!["item"]);
}
#[test] fn null_safe_on_null_object_returns_null() {
    assert_eq!(run_prints(r#"<?php
$obj = null;
$result = $obj?->method();
echo $result ?? 'nil';
"#), vec!["nil"]);
}

// ── Spaceship operator <=> ────────────────────────────────────

#[test] fn spaceship_less_than() {
    assert_eq!(run_prints(r#"<?php echo 1 <=> 2; "#), vec!["-1"]);
}
#[test] fn spaceship_equal() {
    assert_eq!(run_prints(r#"<?php echo 2 <=> 2; "#), vec!["0"]);
}
#[test] fn spaceship_greater_than() {
    assert_eq!(run_prints(r#"<?php echo 3 <=> 2; "#), vec!["1"]);
}
#[test] fn spaceship_in_usort() {
    assert_eq!(run_prints(r#"<?php
$a = [3,1,4,1,5,9,2,6];
usort($a, fn($x,$y) => $x <=> $y);
echo implode(',', $a);
"#), vec!["1,1,2,3,4,5,6,9"]);
}
#[test] fn spaceship_strings() {
    assert_eq!(run_prints(r#"<?php echo 'apple' <=> 'banana'; "#), vec!["-1"]);
}

// ── Elvis operator ?: ─────────────────────────────────────────

#[test] fn elvis_falsy_uses_right() {
    assert_eq!(run_prints(r#"<?php $v = ''; echo $v ?: 'empty'; "#), vec!["empty"]);
}
#[test] fn elvis_truthy_uses_left() {
    assert_eq!(run_prints(r#"<?php $v = 'hello'; echo $v ?: 'empty'; "#), vec!["hello"]);
}
#[test] fn elvis_vs_null_coalescing_difference() {
    assert_eq!(run_prints(r#"<?php
$v = 0;
echo ($v ?: 'falsy') . '|' . ($v ?? 'null');
"#), vec!["falsy|0"]);
}
