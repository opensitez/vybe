use super::helpers::run_prints;

// ── Named arguments with built-in functions ───────────────────

#[test]
fn named_arg_array_slice() {
    assert_eq!(
        run_prints(
            r#"<?php echo implode(',', array_slice(array: [1,2,3,4,5], offset: 1, length: 3)); "#
        ),
        vec!["2,3,4"]
    );
}
#[test]
fn named_arg_implode() {
    assert_eq!(
        run_prints(r#"<?php echo implode(separator: '-', array: ['a','b','c']); "#),
        vec!["a-b-c"]
    );
}
#[test]
fn named_arg_str_pad() {
    assert_eq!(
        run_prints(
            r#"<?php echo str_pad(string: '42', length: 6, pad_string: '0', pad_type: STR_PAD_LEFT); "#
        ),
        vec!["000042"]
    );
}
#[test]
fn named_arg_substr() {
    assert_eq!(
        run_prints(r#"<?php echo substr(string: 'Hello World', offset: 6); "#),
        vec!["World"]
    );
}
#[test]
fn named_arg_str_contains() {
    assert_eq!(
        run_prints(
            r#"<?php echo str_contains(haystack: 'Hello World', needle: 'World') ? 'yes' : 'no'; "#
        ),
        vec!["yes"]
    );
}

// ── Named arguments out of order ──────────────────────────────

#[test]
fn named_arg_out_of_order_user_func() {
    assert_eq!(
        run_prints(
            r#"<?php
function greet(string $name, string $greeting = 'Hello'): string {
    return "$greeting, $name!";
}
echo greet(greeting: 'Hi', name: 'Alice');
"#
        ),
        vec!["Hi, Alice!"]
    );
}
#[test]
fn named_arg_skip_optional() {
    assert_eq!(
        run_prints(
            r#"<?php
function config(string $key, mixed $default = null, bool $required = false): mixed {
    return $required ? $key : ($default ?? $key);
}
echo config(key: 'host', required: true);
"#
        ),
        vec!["host"]
    );
}

// ── Named arguments with variadic ────────────────────────────

#[test]
fn named_arg_before_variadic() {
    assert_eq!(
        run_prints(
            r#"<?php
function tag(string $name, string ...$classes): string {
    return "<$name class=\"" . implode(' ', $classes) . "\">";
}
echo tag(name: 'div', 'foo', 'bar');
"#
        ),
        vec!["<div class=\"foo bar\">"]
    );
}

// ── Named arguments in constructors ──────────────────────────

#[test]
fn named_arg_constructor_out_of_order() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function __construct(public float $x, public float $y, public float $z = 0.0) {}
}
$p = new Point(y: 2.0, x: 1.0);
echo $p->x . ',' . $p->y . ',' . $p->z;
"#
        ),
        vec!["1,2,0"]
    );
}
#[test]
fn named_arg_with_default_skipped() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    public function __construct(
        public string $host = 'localhost',
        public int $port = 3306,
        public string $dbname = 'default'
    ) {}
}
$c = new Config(dbname: 'myapp');
echo $c->host . ':' . $c->port . '/' . $c->dbname;
"#
        ),
        vec!["localhost:3306/myapp"]
    );
}

// ── Named arguments with spread ───────────────────────────────

#[test]
fn named_arg_spread_from_assoc() {
    assert_eq!(
        run_prints(
            r#"<?php
function add(int $a, int $b, int $c): int { return $a + $b + $c; }
$args = ['b' => 2, 'c' => 3, 'a' => 1];
echo add(...$args);
"#
        ),
        vec!["6"]
    );
}

// ── PHP 8.0+ match with named args ───────────────────────────

#[test]
fn named_arg_in_match_arm_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function format(float $n, int $decimals = 2, string $separator = '.'): string {
    return number_format($n, $decimals, $separator, ',');
}
echo format(n: 1234.5678, decimals: 1);
"#
        ),
        vec!["1,234.6"]
    );
}

// ── Named arguments duplicate detection ──────────────────────

#[test]
fn named_arg_error_on_duplicate() {
    assert_eq!(
        run_prints(
            r#"<?php
function test(int $a): int { return $a; }
try { test(a: 1, a: 2); } catch (Error $e) { echo 'duplicate'; }
"#
        ),
        vec!["duplicate"]
    );
}

// ── Named arguments interleaved with positional ───────────────

#[test]
fn named_positional_mixed() {
    assert_eq!(
        run_prints(
            r#"<?php
function box(string $color, int $width, int $height): string {
    return "$color {$width}x$height";
}
echo box('red', height: 10, width: 5);
"#
        ),
        vec!["red 5x10"]
    );
}

// ── Named args with array_map ─────────────────────────────────

#[test]
fn named_arg_array_map_callback() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = array_map(callback: fn($n) => $n * 2, array: [1, 2, 3]);
echo implode(',', $result);
"#
        ),
        vec!["2,4,6"]
    );
}

// ── Named args in static methods ─────────────────────────────

#[test]
fn named_arg_static_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Converter {
    public static function convert(float $value, string $from = 'C', string $to = 'F'): float {
        if ($from === 'C' && $to === 'F') return $value * 9/5 + 32;
        if ($from === 'F' && $to === 'C') return ($value - 32) * 5/9;
        return $value;
    }
}
echo Converter::convert(value: 100.0, to: 'F', from: 'C');
"#
        ),
        vec!["212"]
    );
}
