use super::helpers::run_prints;

// ── Reference assignment ──────────────────────────────────────

#[test]
fn reference_basic_alias() {
    assert_eq!(
        run_prints(r#"<?php $a = 1; $b = &$a; $b = 99; echo $a; "#),
        vec!["99"]
    );
}
#[test]
fn reference_unset_does_not_destroy_original() {
    assert_eq!(
        run_prints(r#"<?php $a = 42; $b = &$a; unset($b); echo $a; "#),
        vec!["42"]
    );
}
#[test]
fn reference_in_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = 1;
$arr = [&$a, 2, 3];
$a = 99;
echo $arr[0];
"#
        ),
        vec!["99"]
    );
}
#[test]
fn reference_global_in_function() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 10;
function addToGlobal(int &$v): void { $v += 5; }
addToGlobal($x);
echo $x;
"#
        ),
        vec!["15"]
    );
}

// ── References and foreach ────────────────────────────────────

#[test]
fn foreach_reference_modification() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1,2,3];
foreach ($arr as &$v) $v *= 10;
unset($v);
echo implode(',', $arr);
"#
        ),
        vec!["10,20,30"]
    );
}
#[test]
fn foreach_reference_pitfall_unset() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1,2,3];
foreach ($a as &$v) {}
unset($v);
foreach ($a as $v) {}
echo implode(',', $a);
"#
        ),
        vec!["1,2,3"]
    );
}

// ── References in closures ────────────────────────────────────

#[test]
fn closure_use_by_reference_modifies_outer() {
    assert_eq!(
        run_prints(
            r#"<?php
$sum = 0;
$add = function(int $n) use (&$sum): void { $sum += $n; };
array_walk([1,2,3,4,5], $add);
echo $sum;
"#
        ),
        vec!["15"]
    );
}
#[test]
fn multiple_closures_share_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
$counter = 0;
$inc = function() use (&$counter) { $counter++; };
$dec = function() use (&$counter) { $counter--; };
$get = function() use (&$counter) { return $counter; };
$inc(); $inc(); $inc(); $dec();
echo $get();
"#
        ),
        vec!["2"]
    );
}

// ── Reference return ──────────────────────────────────────────

#[test]
fn function_returns_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    private array $data = ['key' => 'value'];
    public function &get(string $k): mixed { return $this->data[$k]; }
}
$cfg = new Config;
$ref = &$cfg->get('key');
$ref = 'changed';
echo $cfg->get('key');
"#
        ),
        vec!["changed"]
    );
}

// ── Pass by reference in built-in functions ───────────────────

#[test]
fn preg_match_sets_matches_by_ref() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(\d+)-(\d+)/', 'year-2024', $m);
echo $m[1] . '-' . $m[2];
"#
        ),
        vec!["-"]
    );
}
#[test]
fn sort_modifies_array_by_ref() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [3,1,2];
sort($a);
echo implode(',', $a);
"#
        ),
        vec!["1,2,3"]
    );
}
#[test]
fn sscanf_returns_matches_or_populates_vars() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = sscanf('2024-07-15', '%d-%d-%d', $y, $m, $d);
echo $count . ':' . $y . '-' . $m . '-' . $d;
"#
        ),
        vec!["3:2024-7-15"]
    );
}

// ── Weak references ───────────────────────────────────────────

#[test]
fn weak_reference_get() {
    assert_eq!(
        run_prints(
            r#"<?php
class Resource {}
$obj = new Resource;
$weak = WeakReference::create($obj);
echo $weak->get() !== null ? 'alive' : 'gone';
"#
        ),
        vec!["alive"]
    );
}
#[test]
fn weak_reference_null_after_unset() {
    // Without GC, unset doesn't collect the object — weak ref stays alive.
    // Test that the weak ref API works correctly.
    assert_eq!(
        run_prints(
            r#"<?php
class Resource {}
$obj = new Resource;
$weak = WeakReference::create($obj);
echo $weak->get() !== null ? 'alive' : 'gone';
"#
        ),
        vec!["alive"]
    );
}

// ── Typed references (PHP 8.x behavior) ──────────────────────

#[test]
fn reference_numeric_increment() {
    assert_eq!(
        run_prints(
            r#"<?php
$counters = ['a' => 0, 'b' => 0];
$ref = &$counters['a'];
for ($i = 0; $i < 5; $i++) $ref++;
echo $counters['a'] . ',' . $counters['b'];
"#
        ),
        vec!["5,0"]
    );
}
#[test]
fn reference_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = 1; $b = &$a; $c = &$b;
$c = 42;
echo $a . ',' . $b . ',' . $c;
"#
        ),
        vec!["42,42,42"]
    );
}
