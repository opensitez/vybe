use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Advanced Type System — Union types, Intersection types, DNF types, never, void, mixed, standalone null/false/true
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php80_union_types_argument_and_return() {
    let out = run_prints(
        r#"<?php
function processId(int|string $id): int|string {
    if (is_int($id)) return $id * 2;
    return strtoupper($id);
}

echo processId(10) . " | " . processId("abc");
"#,
    );
    assert_eq!(out, vec!["20 | ABC"]);
}

#[test]
fn test_php81_intersection_types_parameter() {
    let out = run_prints(
        r#"<?php
interface CountableCollection extends Countable, ArrayAccess {}

class CustomCollection implements CountableCollection {
    private array $items = [10, 20];
    public function count(): int { return count($this->items); }
    public function offsetExists($o): bool { return isset($this->items[$o]); }
    public function offsetGet($o): mixed { return $this->items[$o]; }
    public function offsetSet($o, $v): void {}
    public function offsetUnset($o): void {}
}

function inspect(Countable&ArrayAccess $coll): string {
    return "Count=" . count($coll) . " First=" . $coll[0];
}

echo inspect(new CustomCollection());
"#,
    );
    assert_eq!(out, vec!["Count=2 First=10"]);
}

#[test]
fn test_php81_never_return_type_declaration() {
    let out = run_prints(
        r#"<?php
function stopExecution(string $msg): never {
    throw new RuntimeException($msg);
}

try {
    stopExecution("Halted");
} catch (RuntimeException $e) {
    echo "NEVER_RETURN: " . $e->getMessage();
}
"#,
    );
    assert_eq!(out, vec!["NEVER_RETURN: Halted"]);
}

#[test]
fn test_php80_mixed_type_annotation() {
    let out = run_prints(
        r#"<?php
function dumpValue(mixed $val): string {
    return gettype($val);
}

echo dumpValue(123) . " " . dumpValue("hello") . " " . dumpValue([1, 2]);
"#,
    );
    assert_eq!(out, vec!["integer string array"]);
}

#[test]
fn test_php82_standalone_null_false_true_types() {
    compile_ok(
        r#"<?php
function getNull(): null {
    return null;
}

function alwaysFalse(): false {
    return false;
}

function alwaysTrue(): true {
    return true;
}

echo is_null(getNull()) && !alwaysFalse() && alwaysTrue() ? "TYPES_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php82_dnf_disjunctive_normal_form_types() {
    compile_ok(
        r#"<?php
interface HasId { public function getId(): int; }
interface HasName { public function getName(): string; }

class Entity implements HasId, HasName {
    public function getId(): int { return 1; }
    public function getName(): string { return "e1"; }
}

function processEntity((HasId&HasName)|string $target): string {
    if (is_string($target)) return $target;
    return $target->getName();
}

echo processEntity(new Entity()) . " " . processEntity("literal");
"#,
    );
}

#[test]
fn test_php_void_return_type_enforcement() {
    compile_ok(
        r#"<?php
function doWork(): void {
    // Return with no value allowed
    return;
}

doWork();
"#,
    );
}

#[test]
fn test_php_nullable_type_shorthand() {
    compile_ok(
        r#"<?php
function findName(?int $id): ?string {
    if ($id === 1) return "Alice";
    return null;
}

echo findName(1) . " " . (findName(2) ?? "NULL");
"#,
    );
}

#[test]
fn test_php81_type_variance_in_methods() {
    compile_ok(
        r#"<?php
class ParentService {
    public function handle(int|float $num): int|float|string { return $num; }
}

class ChildService extends ParentService {
    // Parameter type widening & Return type narrowing
    public function handle(int|float|string $num): int|float { return 42; }
}

$cs = new ChildService();
echo $cs->handle("123");
"#,
    );
}

#[test]
fn test_php_property_type_hints_typed_properties() {
    compile_ok(
        r#"<?php
class UserProfile {
    public string $name;
    public int $age;
    public ?string $bio = null;
    public array $tags = [];
}

$up = new UserProfile();
$up->name = "Bob";
$up->age = 30;
echo "{$up->name} {$up->age}";
"#,
    );
}
