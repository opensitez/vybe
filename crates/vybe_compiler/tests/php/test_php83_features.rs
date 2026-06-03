use super::helpers::run_prints;

// ── Typed class constants (PHP 8.3) ───────────────────────────

#[test]
fn typed_class_constant_string() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config { const string VERSION = '8.3'; }
echo Config::VERSION;
"#
        ),
        vec!["8.3"]
    );
}
#[test]
fn typed_class_constant_int() {
    assert_eq!(
        run_prints(
            r#"<?php
class Limits { const int MAX = 100; const int MIN = 0; }
echo Limits::MAX - Limits::MIN;
"#
        ),
        vec!["100"]
    );
}
#[test]
fn typed_class_constant_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface HasMax { const int MAX = 255; }
class ColorChannel implements HasMax {}
echo ColorChannel::MAX;
"#
        ),
        vec!["255"]
    );
}

// ── #[Override] attribute (PHP 8.3) ───────────────────────────

#[test]
fn override_attribute_valid() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public function hello(): string { return 'base'; } }
class Child extends Base {
    #[\Override]
    public function hello(): string { return 'child'; }
}
echo (new Child)->hello();
"#
        ),
        vec!["child"]
    );
}

// ── Dynamic class constant access ────────────────────────────

#[test]
fn dynamic_class_constant_access() {
    assert_eq!(
        run_prints(
            r#"<?php
class Status { const ACTIVE = 1; const INACTIVE = 0; }
$name = 'ACTIVE';
echo Status::{$name};
"#
        ),
        vec!["1"]
    );
}

// ── array_sum / array_product with empty ──────────────────────

#[test]
fn array_sum_empty_returns_zero() {
    assert_eq!(run_prints(r#"<?php echo array_sum([]); "#), vec!["0"]);
}
#[test]
fn array_product_empty_returns_one() {
    assert_eq!(run_prints(r#"<?php echo array_product([]); "#), vec!["1"]);
}

// ── str_split / mb_str_split ──────────────────────────────────

#[test]
fn str_split_default_chars() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', str_split('hello')); "#),
        vec!["h,e,l,l,o"]
    );
}
#[test]
fn str_split_with_length() {
    assert_eq!(
        run_prints(r#"<?php echo implode('|', str_split('abcdef', 2)); "#),
        vec!["ab|cd|ef"]
    );
}
#[test]
fn mb_str_split_unicode() {
    assert_eq!(
        run_prints(
            r#"<?php
$chars = mb_str_split('héllo');
echo count($chars) . ':' . $chars[1];
"#
        ),
        vec!["5:é"]
    );
}

// ── Readonly classes (PHP 8.2) ────────────────────────────────

#[test]
fn readonly_class_all_props_readonly() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Point { public function __construct(public float $x, public float $y) {} }
$p = new Point(3.0, 4.0);
echo $p->x . ',' . $p->y;
"#
        ),
        vec!["3,4"]
    );
}
#[test]
fn readonly_class_throws_on_mutation() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class VO { public function __construct(public int $val) {} }
$v = new VO(1);
try { $v->val = 2; } catch (Error $e) { echo 'immutable'; }
"#
        ),
        vec!["immutable"]
    );
}
#[test]
fn readonly_class_cloneable() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class Coord { public function __construct(public float $lat, public float $lon) {} }
$a = new Coord(1.0, 2.0);
$b = clone $a;
echo ($a->lat === $b->lat) ? 'same' : 'diff';
"#
        ),
        vec!["same"]
    );
}

// ── Disjunctive Normal Form (DNF) types (PHP 8.2) ────────────

#[test]
fn dnf_type_nullable_intersection() {
    assert_eq!(
        run_prints(
            r#"<?php
interface A {}
interface B {}
class C implements A, B {}
function test((A&B)|null $obj): string {
    return $obj === null ? 'null' : 'obj';
}
echo test(new C) . ',' . test(null);
"#
        ),
        vec!["obj,null"]
    );
}

// ── str_pad with multibyte safe alternative ───────────────────

#[test]
fn mb_str_pad_php83() {
    assert_eq!(
        run_prints(
            r#"<?php
if (function_exists('mb_str_pad')) {
    echo mb_str_pad('hi', 6, '-', STR_PAD_BOTH);
} else {
    echo str_pad('hi', 6, '-', STR_PAD_BOTH);
}
"#
        ),
        vec!["--hi--"]
    );
}

// ── Fibers improvements (PHP 8.2 / 8.3) ──────────────────────

#[test]
fn fiber_getCurrent_inside_fiber() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    echo Fiber::getCurrent() !== null ? 'inside' : 'outside';
    Fiber::suspend();
});
$fiber->start();
"#
        ),
        vec!["inside"]
    );
}

// ── New string functions PHP 8.3 ─────────────────────────────

#[test]
fn json_validate_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$valid = '{"key":"value"}';
$invalid = '{bad json}';
if (function_exists('json_validate')) {
    echo json_validate($valid) ? 'ok' : 'fail';
    echo json_validate($invalid) ? 'ok' : 'fail';
} else {
    json_decode($valid); $ok1 = json_last_error() === JSON_ERROR_NONE;
    json_decode($invalid); $ok2 = json_last_error() === JSON_ERROR_NONE;
    echo $ok1 ? 'ok' : 'fail';
    echo $ok2 ? 'ok' : 'fail';
}
"#
        ),
        vec!["okfail"]
    );
}
