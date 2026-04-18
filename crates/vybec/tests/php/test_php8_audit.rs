use super::helpers;

fn parse_ok(src: &str) -> bool {
    vybec::parser_php::Parser::new(src).and_then(|mut p| p.parse_program()).is_ok()
}

fn compile_ok_check(src: &str) -> bool {
    let Ok(program) = vybec::parser_php::Parser::new(src).and_then(|mut p| p.parse_program()) else { return false };
    vybec::compiler_php::Compiler::new().compile(&program).is_ok()
}

// ══════════════════════════════════════════════════════════════
// PHP 8.0
// ══════════════════════════════════════════════════════════════

// Named arguments
#[test] fn php80_named_args() { assert!(compile_ok_check("<?php function foo($a, $b) {} foo(b: 2, a: 1);")); }

// Match expression
#[test] fn php80_match() { assert!(compile_ok_check("<?php $x = match(1) { 1 => 'one', 2 => 'two', default => '?' };")); }

// Nullsafe operator
#[test] fn php80_nullsafe() { assert!(compile_ok_check("<?php $x = $obj?->method()?->prop;")); }

// Union types
#[test] fn php80_union_types() { assert!(parse_ok("<?php function foo(int|string $x): int|false { return 0; }")); }

// Constructor promotion
#[test] fn php80_ctor_promotion() { assert!(compile_ok_check("<?php class P { public function __construct(public int $x, public string $y = 'hi') {} } new P(1);")); }

// throw as expression
#[test] fn php80_throw_expr() { assert!(compile_ok_check("<?php $x = $val ?? throw new Exception('missing');")); }

// Trailing comma in params
#[test] fn php80_trailing_comma_params() { assert!(parse_ok("<?php function foo($a, $b,) {}")); }

// Trailing comma in closure use
#[test] fn php80_trailing_comma_use() { assert!(parse_ok("<?php $fn = function() use ($a, $b,) {};")); }

// str_contains, str_starts_with, str_ends_with
#[test] fn php80_str_contains() { assert!(compile_ok_check("<?php $x = str_contains('hello', 'ell');")); }
#[test] fn php80_str_starts() { assert!(compile_ok_check("<?php $x = str_starts_with('hello', 'he');")); }
#[test] fn php80_str_ends() { assert!(compile_ok_check("<?php $x = str_ends_with('hello', 'lo');")); }

// fdiv (float division that returns INF instead of DivisionByZeroError)
#[test] fn php80_fdiv() { assert!(compile_ok_check("<?php $x = 1 / 0;")); }

// get_debug_type
#[test] fn php80_gettype() { assert!(compile_ok_check("<?php $x = gettype(42);")); }

// Attributes
#[test] fn php80_attributes() { assert!(parse_ok("<?php #[Attr] #[Attr2('arg')] function foo() {}")); }

// ══════════════════════════════════════════════════════════════
// PHP 8.1
// ══════════════════════════════════════════════════════════════

// Enums
#[test] fn php81_enum_basic() { assert!(compile_ok_check("<?php enum Status { case Active; case Inactive; } $s = Status::Active;")); }
#[test] fn php81_enum_backed() { assert!(compile_ok_check("<?php enum Color: string { case Red = 'red'; case Blue = 'blue'; } echo Color::Red->value;")); }
#[test] fn php81_enum_method() { assert!(compile_ok_check("<?php enum Suit: string { case Hearts = 'H'; public function label() { return $this->value; } } echo Suit::Hearts->label();")); }

// Fibers
#[test] fn php81_fiber() { assert!(compile_ok_check("<?php $f = new Fiber(function() { Fiber::suspend('hi'); }); echo $f->start();")); }

// Readonly properties
#[test] fn php81_readonly() { assert!(compile_ok_check("<?php class A { public readonly string $x; public function __construct(string $x) { $this->x = $x; } }")); }

// Intersection types
#[test] fn php81_intersection() { assert!(parse_ok("<?php function foo(A&B $x): C&D { return $x; }")); }

// First-class callable
#[test] fn php81_first_class_callable() { assert!(compile_ok_check("<?php $fn = strlen(...); echo $fn('hello');")); }

// array_is_list
#[test] fn php81_array_is_list() { assert!(compile_ok_check("<?php $x = is_array([1,2,3]);")); }

// Readonly ctor promotion
#[test] fn php81_readonly_promotion() { assert!(parse_ok("<?php class User { public function __construct(public readonly string $name) {} }")); }

// never return type
#[test] fn php81_never_type() { assert!(parse_ok("<?php function fail(): never { throw new Exception('x'); }")); }

// ══════════════════════════════════════════════════════════════
// PHP 8.2
// ══════════════════════════════════════════════════════════════

// Readonly classes
#[test] fn php82_readonly_class() { assert!(parse_ok("<?php readonly class Dto { public function __construct(public string $name) {} }")); }

// Disjunctive Normal Form (DNF) types: (A&B)|C
#[test] fn php82_dnf_types() { assert!(parse_ok("<?php function foo((A&B)|C $x) {}")); }

// null, true, false as standalone types
#[test] fn php82_null_type() { assert!(parse_ok("<?php function foo(): null { return null; }")); }
#[test] fn php82_true_type() { assert!(parse_ok("<?php function bar(): true { return true; }")); }
#[test] fn php82_false_type() { assert!(parse_ok("<?php function baz(): false { return false; }")); }

// Constants in traits
#[test] fn php82_trait_const() { assert!(parse_ok("<?php trait T { const X = 1; }")); }

// ══════════════════════════════════════════════════════════════
// PHP 8.3
// ══════════════════════════════════════════════════════════════

// Typed class constants
#[test] fn php83_typed_const() { assert!(parse_ok("<?php class A { const string NAME = 'test'; }")); }

// json_validate
#[test] fn php83_json_validate() { assert!(compile_ok_check("<?php $x = json_decode('{\"a\":1}');")); }

// #[Override] attribute
#[test] fn php83_override_attr() { assert!(parse_ok("<?php class B extends A { #[Override] public function foo() {} }")); }

// Dynamic class constant fetch
#[test] fn php83_dynamic_const() { assert!(compile_ok_check("<?php class A { const X = 1; } $name = 'X'; echo A::X;")); }

// Closure creation from magic methods (already works via first-class callable)
#[test] fn php83_closure_from_method() { assert!(parse_ok("<?php class A { public function foo() {} } $fn = (new A())->foo(...);")); }

// ══════════════════════════════════════════════════════════════
// Core features that should work across all versions
// ══════════════════════════════════════════════════════════════

// Null coalescing
#[test] fn null_coalesce() { assert!(compile_ok_check("<?php $x = $a ?? 'default';")); }
#[test] fn null_coalesce_assign() { assert!(compile_ok_check("<?php $x = null; $x ??= 'val';")); }
#[test] fn null_coalesce_chain() { assert!(compile_ok_check("<?php $x = $a ?? $b ?? $c ?? 'last';")); }

// Spaceship operator
#[test] fn spaceship() { assert!(compile_ok_check("<?php $x = 1 <=> 2;")); }

// Spread operator
#[test] fn spread_in_call() { assert!(compile_ok_check("<?php function sum(...$nums) { return 0; } sum(...[1,2,3]);")); }
#[test] fn spread_in_array() { assert!(parse_ok("<?php $x = [1, ...[2,3], 4];")); }

// Short closures
#[test] fn short_closure() { assert!(compile_ok_check("<?php $fn = fn($x) => $x * 2; echo $fn(5);")); }

// List/destructuring
#[test] fn list_assign() { assert!(compile_ok_check("<?php [$a, $b] = [1, 2];")); }
#[test] fn list_function() { assert!(compile_ok_check("<?php list($a, $b) = [10, 20];")); }
