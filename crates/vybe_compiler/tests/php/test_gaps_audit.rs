use super::helpers::{compile_ok_check, parse_ok};

fn try_compile(src: &str) -> Result<(), String> {
    if compile_ok_check(src) {
        Ok(())
    } else {
        Err("compile failed".into())
    }
}

// ── Complex string interpolation ────────────────────────────
#[test]
fn interp_array_access() {
    assert!(
        parse_ok(r#"<?php $a=[1]; echo "v: {$a[0]}";"#),
        "array interp parse"
    );
}
#[test]
fn interp_property() {
    assert!(
        parse_ok(r#"<?php echo "name: {$obj->name}";"#),
        "prop interp parse"
    );
}

// ── Enum (PHP 8.1) ──────────────────────────────────────────
#[test]
fn enum_parse() {
    assert!(
        parse_ok("<?php enum Color { case Red; case Green; case Blue; }"),
        "enum parse"
    );
}
#[test]
fn enum_backed_parse() {
    assert!(
        parse_ok("<?php enum Suit: string { case Hearts = 'H'; case Diamonds = 'D'; }"),
        "backed enum parse"
    );
}

// ── Type hints (parsed but ignored) ─────────────────────────
#[test]
fn type_hints_parse() {
    assert!(
        parse_ok("<?php function add(int $a, string $b): int { return 0; }"),
        "type hints parse"
    );
}
#[test]
fn union_types_parse() {
    assert!(
        parse_ok("<?php function foo(int|string $x): void {}"),
        "union types parse"
    );
}
#[test]
fn nullable_type_parse() {
    assert!(
        parse_ok("<?php function foo(?int $x): ?string { return null; }"),
        "nullable type parse"
    );
}

// ── Readonly (PHP 8.1) ──────────────────────────────────────
#[test]
fn readonly_parse() {
    assert!(
        parse_ok("<?php class A { public readonly string $name; }"),
        "readonly parse"
    );
}

// ── Abstract class ──────────────────────────────────────────
#[test]
fn abstract_class() {
    assert!(try_compile("<?php abstract class Shape { public $sides; abstract public function area(); } class Circle extends Shape { public function area() { return 3.14; } }").is_ok(), "abstract class compile");
}

// ── Null-safe chaining ──────────────────────────────────────
#[test]
fn nullsafe_chain() {
    assert!(
        try_compile("<?php $x = $a?->b?->c;").is_ok(),
        "nullsafe chain compile"
    );
}

// ── Spread in arrays ────────────────────────────────────────
#[test]
fn spread_array() {
    assert!(
        parse_ok("<?php $x = [...[1,2], ...[3,4]];"),
        "spread array parse"
    );
}

// ── Multi-catch ─────────────────────────────────────────────
#[test]
fn multi_catch() {
    assert!(
        try_compile("<?php try { throw new Exception('x'); } catch (Exception $e) { echo $e; }")
            .is_ok(),
        "multi catch compile"
    );
}

// ── Attributes ──────────────────────────────────────────────
#[test]
fn attributes_parse() {
    assert!(
        parse_ok("<?php #[Attr] function foo() {}"),
        "attributes parse"
    );
}

// ── Named arguments ─────────────────────────────────────────
#[test]
fn named_args() {
    assert!(
        try_compile("<?php function foo($a, $b) {} foo(b: 2, a: 1);").is_ok(),
        "named args compile"
    );
}

// ── First class callable ────────────────────────────────────
#[test]
fn first_class_callable() {
    assert!(
        parse_ok("<?php $fn = strlen(...);"),
        "first class callable parse"
    );
}

// ── Constructor promotion ───────────────────────────────────
#[test]
fn ctor_promotion() {
    assert!(
        parse_ok(
            "<?php class P { public function __construct(public string $name, public int $age) {} }"
        ),
        "ctor promotion parse"
    );
}
