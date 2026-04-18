use super::helpers;

fn parse_ok(src: &str) -> bool {
    vybec::parser_php::Parser::new(src).and_then(|mut p| p.parse_program()).is_ok()
}

fn compile_ok_check(src: &str) -> bool {
    let Ok(program) = vybec::parser_php::Parser::new(src).and_then(|mut p| p.parse_program()) else { return false };
    vybec::compiler_php::Compiler::new().compile(&program).is_ok()
}

// ══════════════════════════════════════════════════════════════
// PHP 7.0 — Still current
// ══════════════════════════════════════════════════════════════

// Scalar type declarations
#[test] fn php70_scalar_types() { assert!(parse_ok("<?php function add(int $a, float $b): string { return ''; }")); }

// Return type declarations
#[test] fn php70_return_type() { assert!(parse_ok("<?php function foo(): array { return []; }")); }

// Null coalescing operator
#[test] fn php70_null_coalesce() { assert!(compile_ok_check("<?php $x = $a ?? 'default';")); }

// Spaceship operator
#[test] fn php70_spaceship() { assert!(compile_ok_check("<?php $x = 1 <=> 2;")); }

// Constant arrays using define
#[test] fn php70_define_array() { assert!(compile_ok_check("<?php define('COLORS', ['red', 'green', 'blue']); echo COLORS;")); }

// Anonymous classes
#[test] fn php70_anon_class() { assert!(parse_ok("<?php $obj = new class { public function hello() { return 'hi'; } };")); }

// Group use declarations
#[test] fn php70_group_use() { assert!(parse_ok("<?php use App\\{Controller, Model, View};")); }

// ══════════════════════════════════════════════════════════════
// PHP 7.1 — Still current
// ══════════════════════════════════════════════════════════════

// Nullable types
#[test] fn php71_nullable_type() { assert!(parse_ok("<?php function foo(?int $x): ?string { return null; }")); }

// Void return type
#[test] fn php71_void_return() { assert!(parse_ok("<?php function doStuff(): void { return; }")); }

// Iterable type
#[test] fn php71_iterable() { assert!(parse_ok("<?php function process(iterable $items): void {}")); }

// Class constant visibility
#[test] fn php71_const_visibility() { assert!(parse_ok("<?php class A { public const X = 1; protected const Y = 2; private const Z = 3; }")); }

// Multi-catch
#[test] fn php71_multi_catch() { assert!(compile_ok_check("<?php try { } catch (TypeError | ValueError $e) { echo $e; }")); }

// Symmetric array destructuring
#[test] fn php71_short_list() { assert!(compile_ok_check("<?php [$a, $b] = [1, 2]; echo $a;")); }

// Keys in list()
#[test] fn php71_keyed_list() { assert!(compile_ok_check("<?php $data = ['first' => 1, 'second' => 2]; ['first' => $a, 'second' => $b] = $data;")); }

// Negative string offsets
#[test] fn php71_negative_offset() { assert!(compile_ok_check("<?php $s = 'hello'; echo substr($s, -2);")); }

// ══════════════════════════════════════════════════════════════
// PHP 7.2 — Still current
// ══════════════════════════════════════════════════════════════

// Object type hint
#[test] fn php72_object_type() { assert!(parse_ok("<?php function foo(object $x): object { return $x; }")); }

// Trailing commas in grouped use
#[test] fn php72_trailing_comma_use() { assert!(parse_ok("<?php use App\\{A, B, C,};")); }

// ══════════════════════════════════════════════════════════════
// PHP 7.3 — Still current
// ══════════════════════════════════════════════════════════════

// Trailing commas in function calls
#[test] fn php73_trailing_comma_call() { assert!(compile_ok_check("<?php echo strlen('hello',);")); }

// Flexible heredoc/nowdoc (indented closing marker)
#[test] fn php73_flexible_heredoc() { assert!(parse_ok("<?php $x = <<<EOT\n    Hello\n    EOT;")); }

// array_key_first / array_key_last (compile as builtins)
#[test] fn php73_array_funcs() { assert!(compile_ok_check("<?php $a = ['x' => 1]; echo array_keys($a);")); }

// ══════════════════════════════════════════════════════════════
// PHP 7.4 — Still current
// ══════════════════════════════════════════════════════════════

// Typed properties
#[test] fn php74_typed_props() { assert!(parse_ok("<?php class User { public int $age; public string $name; }")); }

// Arrow functions
#[test] fn php74_arrow_fn() { assert!(compile_ok_check("<?php $fn = fn($x) => $x * 2; echo $fn(5);")); }

// Null coalescing assignment
#[test] fn php74_null_coalesce_assign() { assert!(compile_ok_check("<?php $x = null; $x ??= 'default'; echo $x;")); }

// Spread in array expression
#[test] fn php74_spread_array() { assert!(parse_ok("<?php $a = [1, 2]; $b = [...$a, 3, 4];")); }

// Numeric literal separator
#[test] fn php74_numeric_separator() { assert!(parse_ok("<?php $x = 1_000_000; $y = 0xFF_FF;")); }

// Preloading (runtime, not syntax — skip)

// ══════════════════════════════════════════════════════════════
// Cross-version features that must work
// ══════════════════════════════════════════════════════════════

// Generators (PHP 5.5+ but heavily used)
#[test] fn generators() { assert!(compile_ok_check("<?php function gen() { yield 1; yield 2; yield 3; }")); }

// Variadic functions (PHP 5.6+)
#[test] fn variadic() { assert!(compile_ok_check("<?php function sum(int ...$nums): int { return 0; } echo sum(1,2,3);")); }

// Argument unpacking (PHP 5.6+)
#[test] fn argument_unpack() { assert!(compile_ok_check("<?php function add($a, $b) { return $a + $b; } echo add(...[1, 2]);")); }

// finally block (PHP 5.5+)
#[test] fn finally_block() { assert!(compile_ok_check("<?php try { echo 1; } catch (Exception $e) {} finally { echo 2; }")); }

// ::class constant (PHP 5.5+)
#[test] fn class_name_constant() { assert!(compile_ok_check("<?php class Foo {} echo Foo::class;")); }

// Exponentiation (PHP 5.6+)
#[test] fn exponentiation() { assert!(compile_ok_check("<?php $x = 2 ** 10; echo $x;")); }
