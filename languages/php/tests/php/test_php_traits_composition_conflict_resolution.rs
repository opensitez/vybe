use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Traits Composition & Conflict Resolution — use TraitA, TraitB, insteadof, as aliasing, trait properties, abstract methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_trait_basic_method_inclusion() {
    let out = run_prints(
        r#"<?php
trait Loggable {
    public function log(string $msg): string {
        return "[LOG] $msg";
    }
}

class User {
    use Loggable;
}

$u = new User();
echo $u->log("User created");
"#,
    );
    assert_eq!(out, vec!["[LOG] User created"]);
}

#[test]
fn test_php_trait_conflict_resolution_insteadof() {
    let out = run_prints(
        r#"<?php
trait TraitA {
    public function speak(): string { return "A"; }
}
trait TraitB {
    public function speak(): string { return "B"; }
}

class Speaker {
    use TraitA, TraitB {
        TraitA::speak insteadof TraitB;
    }
}

$s = new Speaker();
echo $s->speak();
"#,
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn test_php_trait_method_aliasing_and_visibility_change() {
    let out = run_prints(
        r#"<?php
trait Helper {
    private function internalWork(): string { return "done"; }
}

class Worker {
    use Helper {
        internalWork as public publicWork;
    }
}

$w = new Worker();
echo $w->publicWork();
"#,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn test_php_trait_composed_from_multiple_traits() {
    let out = run_prints(
        r#"<?php
trait Timestampable {
    public string $createdAt = "2024-01-01";
}
trait SoftDeletable {
    public bool $isDeleted = false;
}
trait Auditable {
    use Timestampable, SoftDeletable;
}

class Article {
    use Auditable;
}

$a = new Article();
echo $a->createdAt . " deleted=" . ($a->isDeleted ? "1" : "0");
"#,
    );
    assert_eq!(out, vec!["2024-01-01 deleted=0"]);
}

#[test]
fn test_php_abstract_method_in_trait() {
    compile_ok(
        r#"<?php
trait IdentifiableTrait {
    abstract public function getId(): int;
    
    public function getFormattedId(): string {
        return "ID#" . $this->getId();
    }
}

class Invoice {
    use IdentifiableTrait;
    public function getId(): int { return 42; }
}

$inv = new Invoice();
echo $inv->getFormattedId();
"#,
    );
}

#[test]
fn test_php_trait_static_methods_and_properties() {
    compile_ok(
        r#"<?php
trait SingletonTrait {
    private static ?self $instance = null;
    public static function getInstance(): self {
        return self::$instance ??= new self();
    }
}

class AppConfig {
    use SingletonTrait;
}

$app = AppConfig::getInstance();
echo get_class($app);
"#,
    );
}

#[test]
fn test_php_class_method_precedence_over_trait() {
    compile_ok(
        r#"<?php
trait Greeting {
    public function hello() { return "Trait Hello"; }
}

class CustomGreeting {
    use Greeting;
    public function hello() { return "Class Hello"; }
}

$cg = new CustomGreeting();
echo $cg->hello(); // Class method overrides trait method!
"#,
    );
}

#[test]
fn test_php_trait_aliasing_retaining_original_name() {
    compile_ok(
        r#"<?php
trait OutputHelper {
    public function print() { return "Original Print"; }
}

class Printer {
    use OutputHelper {
        print as customPrint;
    }
}

$p = new Printer();
echo $p->print() . " | " . $p->customPrint();
"#,
    );
}

#[test]
fn test_php_trait_property_compatible_definition() {
    compile_ok(
        r#"<?php
trait Configurable {
    public array $options = [];
}

class Settings {
    use Configurable;
    public array $options = []; // Compatible property redeclaration allowed in PHP 8.0+
}

$s = new Settings();
print_r($s->options);
"#,
    );
}

#[test]
fn test_php_trait_parent_class_precedence() {
    compile_ok(
        r#"<?php
class BaseClass {
    public function say() { return "Base"; }
}

trait TraitSay {
    public function say() { return "Trait"; }
}

class ChildClass extends BaseClass {
    use TraitSay;
}

$c = new ChildClass();
echo $c->say(); // Trait method overrides parent class method!
"#,
    );
}
