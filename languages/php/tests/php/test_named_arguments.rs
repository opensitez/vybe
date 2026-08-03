use super::helpers::{compile_ok, run_prints};

// ── Named arguments (PHP 8.0+) ───────────────────────────────────

#[test]
fn named_arg_basic_user_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function greet(string $name, string $greeting = 'Hello'): string {
    return "$greeting, $name!";
}
echo greet(name: 'Alice') . "\n";
echo greet(name: 'Bob', greeting: 'Hi') . "\n";
"#
        ),
        vec!["Hello, Alice!", "Hi, Bob!"]
    );
}

#[test]
fn named_args_all_named() {
    assert_eq!(
        run_prints(
            r#"<?php
function createUser(string $name, int $age, string $role): string {
    return "$name|$age|$role";
}
echo createUser(name: 'Alice', age: 30, role: 'admin') . "\n";
"#
        ),
        vec!["Alice|30|admin"]
    );
}

#[test]
fn named_args_skip_middle_default() {
    assert_eq!(
        run_prints(
            r#"<?php
function padded(string $s, int $width = 10, string $pad = ' ', int $type = STR_PAD_RIGHT): string {
    return str_pad($s, $width, $pad, $type);
}
$r = padded('hi', pad: '*');
echo strlen($r) . "\n";
echo $r[0] . $r[1] . "\n";
"#
        ),
        vec!["10", "hi"]
    );
}

#[test]
fn named_args_out_of_order() {
    assert_eq!(
        run_prints(
            r#"<?php
function box(int $height, int $width, int $depth): string {
    return "{$width}x{$height}x{$depth}";
}
echo box(depth: 5, width: 3, height: 7) . "\n";
"#
        ),
        vec!["3x7x5"]
    );
}

#[test]
fn named_args_in_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function __construct(public float $x, public float $y, public float $z = 0.0) {}
    public function __toString(): string { return "{$this->x},{$this->y},{$this->z}"; }
}
$p = new Point(y: 2.0, x: 1.0);
echo $p . "\n";
"#
        ),
        vec!["1,2,0"]
    );
}

#[test]
fn named_args_constructor_promotion() {
    assert_eq!(
        run_prints(
            r#"<?php
class Product {
    public function __construct(
        public readonly string $name,
        public readonly float $price,
        public readonly int $stock = 0
    ) {}
}
$p = new Product(price: 9.99, name: 'Widget');
echo $p->name . "\n";
echo $p->price . "\n";
echo $p->stock . "\n";
"#
        ),
        vec!["Widget", "9.99", "0"]
    );
}

#[test]
fn named_args_static_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class MathHelper {
    public static function power(float $base, int $exponent = 2): float {
        return pow($base, $exponent);
    }
}
echo MathHelper::power(base: 3.0, exponent: 3) . "\n";
echo MathHelper::power(base: 4.0) . "\n";
"#
        ),
        vec!["27", "16"]
    );
}

#[test]
fn named_args_instance_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Formatter {
    public function format(string $value, int $decimals = 2, string $dec_point = '.', string $thousands_sep = ','): string {
        return number_format((float)$value, $decimals, $dec_point, $thousands_sep);
    }
}
$f = new Formatter();
echo $f->format(value: '1234567.891', decimals: 1) . "\n";
"#
        ),
        vec!["1,234,567.9"]
    );
}

#[test]
fn named_args_mixed_positional_then_named() {
    assert_eq!(
        run_prints(
            r#"<?php
function rangeSum(int $start, int $end, int $step = 1): int {
    $sum = 0;
    for ($i = $start; $i <= $end; $i += $step) $sum += $i;
    return $sum;
}
echo rangeSum(1, 10) . "\n";
echo rangeSum(1, end: 10, step: 2) . "\n";
"#
        ),
        vec!["55", "25"]
    );
}

#[test]
fn named_arg_with_null_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function coalesce(?string $value, string $default = 'fallback'): string {
    return $value ?? $default;
}
echo coalesce(value: null) . "\n";
echo coalesce(value: 'provided') . "\n";
"#
        ),
        vec!["fallback", "provided"]
    );
}

#[test]
fn named_arg_with_array_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function sumArray(array $items, int $initial = 0): int {
    return array_reduce($items, fn($c, $v) => $c + $v, $initial);
}
echo sumArray(items: [1, 2, 3, 4]) . "\n";
echo sumArray(items: [10, 20], initial: 5) . "\n";
"#
        ),
        vec!["10", "35"]
    );
}

#[test]
fn named_arg_with_closure_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function applyTransform(array $data, callable $transform): array {
    return array_map($transform, $data);
}
$result = applyTransform(data: [1, 2, 3], transform: fn($x) => $x * 3);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["3,6,9"]
    );
}

#[test]
fn named_arg_in_recursive_call() {
    assert_eq!(
        run_prints(
            r#"<?php
function countdown(int $from, int $step = 1): void {
    if ($from <= 0) { echo "done\n"; return; }
    echo $from . "\n";
    countdown(from: $from - $step, step: $step);
}
countdown(from: 3);
"#
        ),
        vec!["3", "2", "1", "done"]
    );
}

#[test]
fn named_args_builtin_array_slice() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [10, 20, 30, 40, 50];
$slice = array_slice(array: $arr, offset: 1, length: 3);
echo implode(',', $slice) . "\n";
"#
        ),
        vec!["20,30,40"]
    );
}

#[test]
fn named_args_builtin_str_pad() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = str_pad(string: 'hi', length: 8, pad_string: '-', pad_type: STR_PAD_BOTH);
echo $r . "\n";
"#
        ),
        vec!["---hi---"]
    );
}

#[test]
fn named_args_builtin_implode() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = implode(separator: ', ', array: ['a', 'b', 'c']);
echo $result . "\n";
"#
        ),
        vec!["a, b, c"]
    );
}

#[test]
fn named_args_builtin_substr() {
    assert_eq!(
        run_prints(
            r#"<?php
echo substr(string: 'Hello World', offset: 6, length: 5) . "\n";
echo substr(string: 'Hello World', offset: 6) . "\n";
"#
        ),
        vec!["World", "World"]
    );
}

#[test]
fn named_args_builtin_in_array_strict() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [1, 2, 3, '4', 5];
echo in_array(needle: 4, haystack: $arr, strict: true) ? 'found' : 'not found';
echo "\n";
echo in_array(needle: 4, haystack: $arr, strict: false) ? 'found' : 'not found';
echo "\n";
"#
        ),
        vec!["not found", "found"]
    );
}

#[test]
fn named_args_builtin_array_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$map = ['a' => 1, 'b' => 2, 'c' => 1, 'd' => 3];
$keys = array_keys(array: $map, filter_value: 1);
echo implode(',', $keys) . "\n";
"#
        ),
        vec!["a,c"]
    );
}

#[test]
fn named_args_combined_with_spread() {
    assert_eq!(
        run_prints(
            r#"<?php
function sumThree(int $a, int $b, int $c): int { return $a + $b + $c; }
$extra = [2, 3];
echo sumThree(a: 1, ...$extra) . "\n";
"#
        ),
        vec!["6"]
    );
}

#[test]
fn named_args_passed_through_wrapper() {
    assert_eq!(
        run_prints(
            r#"<?php
function inner(string $prefix, string $suffix, string $sep = '-'): string {
    return $prefix . $sep . $suffix;
}
function outer(string $prefix, string $suffix, string $sep = '-'): string {
    return inner(prefix: $prefix, suffix: $suffix, sep: $sep);
}
echo outer(prefix: 'foo', suffix: 'bar') . "\n";
echo outer(suffix: 'baz', prefix: 'qux', sep: ':') . "\n";
"#
        ),
        vec!["foo-bar", "qux:baz"]
    );
}

#[test]
fn named_args_with_variadic_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function tagged(string $tag, string ...$items): string {
    return "<$tag>" . implode("</$tag><$tag>", $items) . "</$tag>";
}
echo tagged(tag: 'li', ...[' one', 'two', 'three']) . "\n";
"#
        ),
        vec!["<li> one</li><li>two</li><li>three</li>"]
    );
}

#[test]
fn named_args_in_match_arm_calls() {
    assert_eq!(
        run_prints(
            r#"<?php
function clamp(int $val, int $min, int $max): int { return max($min, min($max, $val)); }
$values = [-5, 5, 15];
foreach ($values as $v) {
    $result = match(true) {
        $v < 0  => clamp(val: $v, min: 0, max: 10),
        $v > 10 => clamp(val: $v, min: 0, max: 10),
        default => $v };
    echo $result . "\n";
}
"#
        ),
        vec!["0", "5", "10"]
    );
}

#[test]
fn named_args_in_arrow_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function repeat(string $s, int $times = 1): string { return str_repeat($s, $times); }
$doubler = fn(string $s) => repeat(s: $s, times: 2);
echo $doubler('ab') . "\n";
echo $doubler('xyz') . "\n";
"#
        ),
        vec!["abab", "xyzxyz"]
    );
}

#[test]
fn named_args_with_interface_method() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Converter {
    public function convert(float $value, string $from, string $to): float;
}
class TempConverter implements Converter {
    public function convert(float $value, string $from, string $to): float {
        if ($from === 'C' && $to === 'F') return $value * 9/5 + 32;
        if ($from === 'F' && $to === 'C') return ($value - 32) * 5/9;
        return $value;
    }
}
$c = new TempConverter();
echo $c->convert(value: 100.0, from: 'C', to: 'F') . "\n";
echo $c->convert(value: 32.0, from: 'F', to: 'C') . "\n";
"#
        ),
        vec!["212", "0"]
    );
}

#[test]
fn named_args_with_abstract_method() {
    compile_ok(
        r#"<?php
abstract class Shape {
    abstract public function area(float $scale = 1.0): float;
}
class Circle extends Shape {
    public function __construct(private float $radius) {}
    public function area(float $scale = 1.0): float {
        return M_PI * $this->radius ** 2 * $scale;
    }
}
$c = new Circle(radius: 5.0);
echo round($c->area(scale: 2.0), 2);
"#,
    );
}

#[test]
fn named_args_with_trait_method() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Formattable {
    public function format(string $template, string $locale = 'en'): string {
        return str_replace('{locale}', $locale, str_replace('{val}', (string)$this->value, $template));
    }
}
class Temperature {
    use Formattable;
    public function __construct(public float $value) {}
}
$t = new Temperature(value: 23.5);
echo $t->format(template: '{val} ({locale})', locale: 'en') . "\n";
"#
        ),
        vec!["23.5 (en)"]
    );
}

#[test]
fn named_args_in_call_user_func_array() {
    assert_eq!(
        run_prints(
            r#"<?php
function buildTag(string $tag, string $content, string $class = ''): string {
    $cls = $class ? " class=\"$class\"" : '';
    return "<$tag$cls>$content</$tag>";
}
$result = call_user_func_array('buildTag', ['tag' => 'div', 'content' => 'Hello', 'class' => 'greeting']);
echo $result . "\n";
"#
        ),
        vec!["<div class=\"greeting\">Hello</div>"]
    );
}

#[test]
fn named_args_in_nested_function_calls() {
    assert_eq!(
        run_prints(
            r#"<?php
function wrap(string $inner, string $outer): string {
    return "<$outer>$inner</$outer>";
}
function buildHtml(string $text, string $inner_tag = 'span', string $outer_tag = 'div'): string {
    return wrap(inner: "<$inner_tag>$text</$inner_tag>", outer: $outer_tag);
}
echo buildHtml(text: 'hello', inner_tag: 'b') . "\n";
"#
        ),
        vec!["<div><b>hello</b></div>"]
    );
}
