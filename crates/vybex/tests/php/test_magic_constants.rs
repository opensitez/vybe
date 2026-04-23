use super::helpers::compile_ok;

// ── __LINE__ ──────────────────────────────────────────────────

#[test] fn magic_line_basic() {
    compile_ok(r#"<?php
$line = __LINE__;
echo is_int($line) ? 'is int' : 'not int';
echo $line > 0 ? ':positive' : ':zero';
"#);
}

#[test] fn magic_line_in_function() {
    compile_ok(r#"<?php
function getLine(): int { return __LINE__; }
$l = getLine();
echo $l > 0 ? 'ok' : 'fail';
"#);
}

#[test] fn magic_line_changes_per_line() {
    compile_ok(r#"<?php
$a = __LINE__;
$b = __LINE__;
$c = __LINE__;
echo ($b > $a && $c > $b) ? 'increasing' : 'fail';
"#);
}

// ── __FILE__ ──────────────────────────────────────────────────

#[test] fn magic_file_basic() {
    compile_ok(r#"<?php
$file = __FILE__;
echo is_string($file) ? 'is string' : 'not string';
"#);
}

#[test] fn magic_file_in_function() {
    compile_ok(r#"<?php
function getFile(): string { return __FILE__; }
echo is_string(getFile()) ? 'ok' : 'fail';
"#);
}

// ── __DIR__ ───────────────────────────────────────────────────

#[test] fn magic_dir_basic() {
    compile_ok(r#"<?php
$dir = __DIR__;
echo is_string($dir) ? 'is string' : 'not string';
"#);
}

#[test] fn magic_dir_vs_file() {
    compile_ok(r#"<?php
$dir  = __DIR__;
$file = __FILE__;
echo strlen($dir) <= strlen($file) ? 'dir shorter or equal' : 'dir longer';
"#);
}

// ── __FUNCTION__ ──────────────────────────────────────────────

#[test] fn magic_function_basic() {
    compile_ok(r#"<?php
function myFunc(): string { return __FUNCTION__; }
echo myFunc();
"#);
}

#[test] fn magic_function_nested() {
    compile_ok(r#"<?php
function outer(): string {
    function inner(): string { return __FUNCTION__; }
    return __FUNCTION__ . ':' . inner();
}
echo outer();
"#);
}

#[test] fn magic_function_closure() {
    compile_ok(r#"<?php
$fn = function(): string { return __FUNCTION__; };
echo $fn();  // "{closure}"
"#);
}

#[test] fn magic_function_arrow() {
    compile_ok(r#"<?php
$fn = fn() => __FUNCTION__;
$result = $fn();
echo is_string($result) ? 'string' : 'fail';
"#);
}

#[test] fn magic_function_global_scope() {
    compile_ok(r#"<?php
// At global scope __FUNCTION__ is empty string
$f = __FUNCTION__;
echo $f === '' ? 'empty at global' : "has value: $f";
"#);
}

// ── __CLASS__ ─────────────────────────────────────────────────

#[test] fn magic_class_basic() {
    compile_ok(r#"<?php
class MyClass {
    public function getClass(): string { return __CLASS__; }
}
echo (new MyClass())->getClass();
"#);
}

#[test] fn magic_class_in_static() {
    compile_ok(r#"<?php
class Foo {
    public static function name(): string { return __CLASS__; }
}
echo Foo::name();
"#);
}

#[test] fn magic_class_in_parent_inherited() {
    compile_ok(r#"<?php
class Base {
    public function whoAmI(): string { return __CLASS__; }
}
class Child extends Base {}
$c = new Child();
echo $c->whoAmI(); // "Base" — __CLASS__ resolves at definition time
"#);
}

#[test] fn magic_class_trait() {
    compile_ok(r#"<?php
trait Identified {
    public function identify(): string { return __CLASS__; }
}
class Alpha { use Identified; }
class Beta  { use Identified; }
echo (new Alpha())->identify();
echo (new Beta())->identify();
"#);
}

// ── __TRAIT__ ─────────────────────────────────────────────────

#[test] fn magic_trait_basic() {
    compile_ok(r#"<?php
trait MyTrait {
    public function traitName(): string { return __TRAIT__; }
}
class UsesTrait { use MyTrait; }
echo (new UsesTrait())->traitName();
"#);
}

#[test] fn magic_trait_multiple_classes() {
    compile_ok(r#"<?php
trait Logging {
    public function source(): string { return __TRAIT__; }
}
class A { use Logging; }
class B { use Logging; }
echo (new A())->source() . ':' . (new B())->source();
"#);
}

#[test] fn magic_trait_outside_trait() {
    compile_ok(r#"<?php
class Plain {
    public function trait(): string { return __TRAIT__; }
}
echo (new Plain())->trait() === '' ? 'empty outside trait' : 'has value';
"#);
}

// ── __METHOD__ ────────────────────────────────────────────────

#[test] fn magic_method_basic() {
    compile_ok(r#"<?php
class Calculator {
    public function add(): string { return __METHOD__; }
}
echo (new Calculator())->add();
"#);
}

#[test] fn magic_method_static() {
    compile_ok(r#"<?php
class Factory {
    public static function create(): string { return __METHOD__; }
}
echo Factory::create();
"#);
}

#[test] fn magic_method_vs_function() {
    compile_ok(r#"<?php
class Util {
    public function run(): string { return __METHOD__; }
}
function standalone(): string { return __FUNCTION__; }
echo (new Util())->run();
echo ':';
echo standalone();
"#);
}

// ── __NAMESPACE__ ─────────────────────────────────────────────

#[test] fn magic_namespace_global() {
    compile_ok(r#"<?php
echo __NAMESPACE__ === '' ? 'global namespace' : __NAMESPACE__;
"#);
}

#[test] fn magic_namespace_declared() {
    compile_ok(r#"<?php
namespace App\Services;
echo __NAMESPACE__;
"#);
}

#[test] fn magic_namespace_in_function() {
    compile_ok(r#"<?php
namespace Domain\Models;
function currentNs(): string { return __NAMESPACE__; }
echo currentNs();
"#);
}

// ── ::class constant ─────────────────────────────────────────

#[test] fn class_constant_basic() {
    compile_ok(r#"<?php
class Foo {}
echo Foo::class;
"#);
}

#[test] fn class_constant_namespaced() {
    compile_ok(r#"<?php
namespace Http;
class Request {}
echo Request::class; // Http\Request
"#);
}

#[test] fn class_constant_on_object() {
    compile_ok(r#"<?php
class Widget { public string $type = 'button'; }
$w = new Widget();
echo $w::class;
"#);
}

#[test] fn class_constant_interface() {
    compile_ok(r#"<?php
interface Drawable { public function draw(): void; }
echo Drawable::class;
"#);
}

#[test] fn class_constant_enum() {
    compile_ok(r#"<?php
enum Color { case Red; case Blue; }
echo Color::class;
"#);
}

// ── Combined use ──────────────────────────────────────────────

#[test] fn magic_all_in_class() {
    compile_ok(r#"<?php
class Diagnostics {
    public function report(): array {
        return [
            'class'    => __CLASS__,
            'method'   => __METHOD__,
            'line'     => __LINE__,
            'file_set' => __FILE__ !== '',
        ];
    }
}
$info = (new Diagnostics())->report();
echo $info['class'] . ':' . $info['method'];
echo ':line=' . ($info['line'] > 0 ? 'ok' : 'fail');
echo ':file=' . ($info['file_set'] ? 'ok' : 'fail');
"#);
}
