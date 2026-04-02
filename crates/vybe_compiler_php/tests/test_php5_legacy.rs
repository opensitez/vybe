mod helpers;

fn parse_ok(src: &str) -> bool {
    vybe_parser_php::Parser::new(src).and_then(|mut p| p.parse_program()).is_ok()
}

fn compile_ok_check(src: &str) -> bool {
    let Ok(program) = vybe_parser_php::Parser::new(src).and_then(|mut p| p.parse_program()) else { return false };
    vybe_compiler_php::Compiler::new().compile(&program).is_ok()
}

// ══════════════════════════════════════════════════════════════
// PHP 5.0 — Core OOP (still fundamental in PHP 8)
// ══════════════════════════════════════════════════════════════

// Visibility modifiers
#[test] fn visibility_public() { assert!(compile_ok_check("<?php class A { public $x; public function foo() {} }")); }
#[test] fn visibility_private() { assert!(compile_ok_check("<?php class A { private $x; private function foo() {} }")); }
#[test] fn visibility_protected() { assert!(compile_ok_check("<?php class A { protected $x; protected function foo() {} }")); }

// Abstract classes and methods
#[test] fn abstract_class() { assert!(compile_ok_check("<?php abstract class Shape { abstract public function area(); } class Circle extends Shape { public function area() { return 3.14; } }")); }

// Interfaces
#[test] fn interface_basic() { assert!(compile_ok_check("<?php interface Loggable { public function log($msg); } class FileLogger implements Loggable { public function log($msg) { echo $msg; } } $l = new FileLogger(); $l->log('hi');")); }

// Final classes and methods
#[test] fn final_class() { assert!(parse_ok("<?php final class Singleton { public static function instance() {} }")); }
#[test] fn final_method() { assert!(parse_ok("<?php class Base { final public function id() { return 1; } }")); }

// Static methods and properties
#[test] fn static_method() { assert!(compile_ok_check("<?php class Counter { public static $count = 0; public static function increment() { Counter::$count = Counter::$count + 1; } } Counter::increment(); echo Counter::$count;")); }

// Type hinting (class/interface names as param types)
#[test] fn class_type_hint() { assert!(parse_ok("<?php class Dog {} function walk(Dog $dog) {}")); }

// __construct / __destruct
#[test] fn constructor() { assert!(compile_ok_check("<?php class A { public $x; public function __construct($x) { $this->x = $x; } } $a = new A(42); echo $a->x;")); }

// Clone
#[test] fn clone_keyword() { assert!(parse_ok("<?php class A { public $x = 1; } $a = new A(); $b = clone $a;")); }

// ══════════════════════════════════════════════════════════════
// PHP 5.1 — PDO, autoload (runtime features, not syntax)
// ══════════════════════════════════════════════════════════════

// ══════════════════════════════════════════════════════════════
// PHP 5.3 — Namespaces, closures, __invoke, goto, nowdoc
// ══════════════════════════════════════════════════════════════

// Namespaces (parse only — no runtime effect in VM)
#[test] fn namespace_decl() { assert!(parse_ok("<?php namespace App\\Models; class User {}")); }
#[test] fn namespace_use() { assert!(parse_ok("<?php use App\\Models\\User; use App\\Services\\{AuthService, MailService};")); }

// Closures (PHP 5.3)
#[test] fn closure_basic() { assert!(compile_ok_check("<?php $greet = function($name) { return 'Hello ' . $name; }; echo $greet('World');")); }
#[test] fn closure_use() { assert!(compile_ok_check("<?php $prefix = 'Mr.'; $fn = function($name) use ($prefix) { return $prefix . ' ' . $name; }; echo $fn('Smith');")); }

// Nowdoc
#[test] fn nowdoc() { assert!(parse_ok("<?php $x = <<<'EOT'\nNo $interpolation here\nEOT;")); }

// Ternary shorthand (Elvis) — $a ?: $b
#[test] fn elvis() { assert!(compile_ok_check("<?php $x = '' ?: 'default'; echo $x;")); }

// const keyword outside class
#[test] fn const_global() { assert!(compile_ok_check("<?php const PI = 3.14159; echo PI;")); }

// ══════════════════════════════════════════════════════════════
// PHP 5.4 — Traits, short array syntax, callable type
// ══════════════════════════════════════════════════════════════

// Traits
#[test] fn trait_basic() { assert!(compile_ok_check(r#"<?php
trait Timestampable {
    public function getCreated() { return $this->created; }
}
class Post { use Timestampable; public $created = '2024-01-01'; }
$p = new Post();
echo $p->getCreated();
"#)); }

// Short array syntax
#[test] fn short_array() { assert!(compile_ok_check("<?php $a = [1, 2, 3]; echo $a[0];")); }

// Function array dereferencing
#[test] fn func_array_deref() { assert!(compile_ok_check("<?php function getArr() { return [1, 2, 3]; } echo getArr()[1];")); }

// Callable type hint
#[test] fn callable_hint() { assert!(parse_ok("<?php function apply(callable $fn, $val) { return $fn($val); }")); }

// $this in closures
#[test] fn closure_this() { assert!(compile_ok_check(r#"<?php
class Foo {
    public $x = 42;
    public function getClosure() {
        return function() { return $this->x; };
    }
}
$f = new Foo();
$fn = $f->getClosure();
"#)); }

// ══════════════════════════════════════════════════════════════
// PHP 5.5 — Generators, finally, ::class, foreach list()
// ══════════════════════════════════════════════════════════════

// Generators
#[test] fn generator_basic() { assert!(compile_ok_check("<?php function gen() { yield 1; yield 2; yield 3; }")); }
#[test] fn generator_keys() { assert!(compile_ok_check("<?php function pairs() { yield 'a'; yield 'b'; }")); }

// finally
#[test] fn try_finally() { assert!(compile_ok_check("<?php try { echo 1; } catch (Exception $e) { echo 2; } finally { echo 3; }")); }

// ::class
#[test] fn class_constant() { assert!(compile_ok_check("<?php class Foo {} echo Foo::class;")); }

// empty() with expressions
#[test] fn empty_expr() { assert!(compile_ok_check("<?php echo empty(trim('  '));")); }

// ══════════════════════════════════════════════════════════════
// PHP 5.6 — Variadic, argument unpacking, const expressions
// ══════════════════════════════════════════════════════════════

// Variadic functions
#[test] fn variadic_func() { assert!(compile_ok_check("<?php function sum(...$nums) { return array_sum($nums); } echo sum(1, 2, 3);")); }

// Argument unpacking
#[test] fn arg_unpack() { assert!(compile_ok_check("<?php function add($a, $b) { return $a + $b; } echo add(...[3, 4]);")); }

// Exponentiation **
#[test] fn power_op() { assert!(compile_ok_check("<?php echo 2 ** 10;")); }

// use function / use const
#[test] fn use_function() { assert!(parse_ok("<?php use function App\\Helpers\\format_date; use const App\\Config\\VERSION;")); }

// Constant scalar expressions
#[test] fn const_expr() { assert!(compile_ok_check("<?php const DOUBLE_PI = 3.14159 * 2;")); }

// ══════════════════════════════════════════════════════════════
// Core PHP features that should always work
// ══════════════════════════════════════════════════════════════

// Multiple interfaces
#[test] fn multi_interface() { assert!(parse_ok("<?php interface A {} interface B {} class C implements A, B {}")); }

// Method chaining
#[test] fn method_chain() { assert!(compile_ok_check("<?php class Q { public function a() { return $this; } public function b() { return $this; } } $q = new Q(); $q->a()->b();")); }

// Nested method calls
#[test] fn nested_calls() { assert!(compile_ok_check("<?php echo strlen(strtoupper(trim('  hello  ')));")); }

// Complex property access
#[test] fn deep_property() { assert!(compile_ok_check("<?php class A { public $b; } class B { public $c = 42; } $a = new A(); $a->b = new B(); echo $a->b->c;")); }

// Array of objects
#[test] fn array_of_objects() { assert!(compile_ok_check(r#"<?php
class Item { public $name; public function __construct($n) { $this->name = $n; } }
$items = [new Item('a'), new Item('b'), new Item('c')];
foreach ($items as $item) { echo $item->name; }
"#)); }

// String functions used everywhere
#[test] fn common_string_ops() { assert!(compile_ok_check(r#"<?php
$email = '  USER@EXAMPLE.COM  ';
$clean = strtolower(trim($email));
$parts = explode('@', $clean);
$user = $parts[0];
$domain = $parts[1];
echo $user . ' at ' . $domain;
"#)); }

// Array manipulation pipeline
#[test] fn array_pipeline() { assert!(compile_ok_check(r#"<?php
$data = ['banana', 'apple', 'cherry', 'date'];
sort($data);
$upper = array_map(fn($s) => strtoupper($s), $data);
$filtered = array_filter($upper, fn($s) => strlen($s) > 4);
echo implode(', ', $filtered);
"#)); }

// Recursive data structure
#[test] fn recursive_structure() { assert!(compile_ok_check(r#"<?php
function flatten($arr) {
    $result = [];
    foreach ($arr as $item) {
        if (is_array($item)) {
            $sub = flatten($item);
            $result = array_merge($result, $sub);
        } else {
            array_push($result, $item);
        }
    }
    return $result;
}
$nested = [1, [2, 3], [4, [5, 6]]];
$flat = flatten($nested);
"#)); }

// Multiple return types via union (conceptual — returns different types)
#[test] fn dynamic_return() { assert!(compile_ok_check(r#"<?php
function parse($input) {
    if (is_numeric($input)) return intval($input);
    if ($input === 'true') return true;
    if ($input === 'null') return null;
    return $input;
}
echo parse('42');
echo parse('hello');
"#)); }

// Static factory + fluent builder
#[test] fn fluent_builder() { assert!(compile_ok_check(r#"<?php
class Response {
    public $status = 200;
    public $body = '';
    public $headers = [];
    public static function create() { return new Response(); }
    public function status($code) { $this->status = $code; return $this; }
    public function body($content) { $this->body = $content; return $this; }
    public function header($key, $val) { array_push($this->headers, $key . ': ' . $val); return $this; }
}
$resp = Response::create()
    ->status(200)
    ->body('{"ok":true}')
    ->header('Content-Type', 'application/json');
echo $resp->body;
"#)); }
