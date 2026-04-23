use super::helpers::compile_ok;

// ── Variable variables ────────────────────────────────────────

#[test] fn variable_variable_basic() {
    compile_ok(r#"<?php
$varName = 'greeting';
$$varName = 'Hello';
echo $greeting;
"#);
}

#[test] fn variable_variable_assign_then_read() {
    compile_ok(r#"<?php
$key = 'color';
$$key = 'blue';
echo $$key;
"#);
}

#[test] fn variable_variable_loop() {
    compile_ok(r#"<?php
$vars = ['a', 'b', 'c'];
foreach ($vars as $i => $name) {
    $$name = $i * 10;
}
echo $a . ',' . $b . ',' . $c;
"#);
}

#[test] fn variable_variable_in_array() {
    compile_ok(r#"<?php
$fields = ['name', 'age', 'city'];
$name = 'Alice';
$age  = 30;
$city = 'NY';
$out = [];
foreach ($fields as $f) { $out[] = $$f; }
echo implode(',', $out);
"#);
}

#[test] fn variable_variable_expression() {
    compile_ok(r#"<?php
$prefix = 'my';
$suffix = 'Var';
$varName = $prefix . $suffix;
$$varName = 42;
echo $myVar;
"#);
}

// ── Dynamic property access ───────────────────────────────────

#[test] fn dynamic_property_get() {
    compile_ok(r#"<?php
class Point { public int $x = 1; public int $y = 2; }
$p = new Point();
$prop = 'x';
echo $p->$prop;
"#);
}

#[test] fn dynamic_property_set() {
    compile_ok(r#"<?php
class Box { public int $width = 0; public int $height = 0; }
$b = new Box();
foreach (['width' => 10, 'height' => 5] as $prop => $val) {
    $b->$prop = $val;
}
echo $b->width . 'x' . $b->height;
"#);
}

#[test] fn dynamic_property_computed() {
    compile_ok(r#"<?php
class Record {
    public string $field1 = 'a';
    public string $field2 = 'b';
    public string $field3 = 'c';
}
$r = new Record();
$result = '';
for ($i = 1; $i <= 3; $i++) {
    $result .= $r->{"field$i"};
}
echo $result;
"#);
}

// ── Dynamic method calls ──────────────────────────────────────

#[test] fn dynamic_method_call() {
    compile_ok(r#"<?php
class Greeter {
    public function hello(): string { return "hello"; }
    public function goodbye(): string { return "goodbye"; }
}
$g = new Greeter();
$method = 'hello';
echo $g->$method();
"#);
}

#[test] fn dynamic_method_dispatch() {
    compile_ok(r#"<?php
class Math {
    public function double(int $n): int { return $n * 2; }
    public function triple(int $n): int { return $n * 3; }
    public function square(int $n): int { return $n * $n; }
}
$m = new Math();
foreach (['double', 'triple', 'square'] as $op) {
    echo $m->$op(4) . ' ';
}
"#);
}

#[test] fn dynamic_static_method() {
    compile_ok(r#"<?php
class Factory {
    public static function makeString(): string { return 'str'; }
    public static function makeInt(): int { return 42; }
}
$method = 'makeString';
echo Factory::$method();
"#);
}

// ── Variable functions ────────────────────────────────────────

#[test] fn variable_function_basic() {
    compile_ok(r#"<?php
function greet(string $name): string { return "Hi, $name!"; }
$fn = 'greet';
echo $fn('Alice');
"#);
}

#[test] fn variable_function_builtin() {
    compile_ok(r#"<?php
$fn = 'strtoupper';
echo $fn('hello');
"#);
}

#[test] fn variable_function_array() {
    compile_ok(r#"<?php
function square(int $n): int { return $n * $n; }
function cube(int $n): int   { return $n * $n * $n; }
$ops = ['square', 'cube'];
foreach ($ops as $fn) { echo $fn(3) . ' '; }
"#);
}

#[test] fn variable_function_conditional() {
    compile_ok(r#"<?php
function add(int $a, int $b): int { return $a + $b; }
function sub(int $a, int $b): int { return $a - $b; }
$mode = 'add';
$fn = $mode;
echo $fn(10, 3);
"#);
}

// ── call_user_func / call_user_func_array ─────────────────────

#[test] fn call_user_func_basic() {
    compile_ok(r#"<?php
function double(int $n): int { return $n * 2; }
$result = call_user_func('double', 21);
echo $result;
"#);
}

#[test] fn call_user_func_closure() {
    compile_ok(r#"<?php
$mul = fn(int $a, int $b) => $a * $b;
echo call_user_func($mul, 6, 7);
"#);
}

#[test] fn call_user_func_method() {
    compile_ok(r#"<?php
class Calc {
    public function add(int $a, int $b): int { return $a + $b; }
}
$c = new Calc();
echo call_user_func([$c, 'add'], 10, 32);
"#);
}

#[test] fn call_user_func_static_method() {
    compile_ok(r#"<?php
class MathUtil {
    public static function square(int $n): int { return $n * $n; }
}
echo call_user_func(['MathUtil', 'square'], 9);
"#);
}

#[test] fn call_user_func_array_basic() {
    compile_ok(r#"<?php
function sum3(int $a, int $b, int $c): int { return $a + $b + $c; }
echo call_user_func_array('sum3', [1, 2, 3]);
"#);
}

#[test] fn call_user_func_array_spread() {
    compile_ok(r#"<?php
$args = [10, 20, 30];
$fn = fn(int ...$nums) => array_sum($nums);
echo call_user_func_array($fn, $args);
"#);
}

// ── Dynamic class instantiation ───────────────────────────────

#[test] fn dynamic_class_new() {
    compile_ok(r#"<?php
class Dog  { public function speak(): string { return "Woof"; } }
class Cat  { public function speak(): string { return "Meow"; } }
class Bird { public function speak(): string { return "Tweet"; } }
foreach (['Dog', 'Cat', 'Bird'] as $cls) {
    $obj = new $cls();
    echo $obj->speak() . ' ';
}
"#);
}

#[test] fn dynamic_class_with_args() {
    compile_ok(r#"<?php
class Color {
    public function __construct(private string $name, private string $hex) {}
    public function __toString(): string { return "{$this->name}:{$this->hex}"; }
}
$className = 'Color';
$obj = new $className('red', '#FF0000');
echo $obj;
"#);
}

// ── Dynamic constant access ───────────────────────────────────

#[test] fn dynamic_constant_fetch() {
    compile_ok(r#"<?php
class Direction {
    const NORTH = 'N';
    const SOUTH = 'S';
    const EAST  = 'E';
    const WEST  = 'W';
}
foreach (['NORTH', 'SOUTH', 'EAST', 'WEST'] as $dir) {
    echo Direction::$$dir ?? constant("Direction::$dir");
}
"#);
}

#[test] fn dynamic_class_constant() {
    compile_ok(r#"<?php
class Status {
    const OK    = 200;
    const ERROR = 500;
}
$const = 'OK';
echo constant("Status::$const");
"#);
}
