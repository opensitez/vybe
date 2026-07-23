use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: OOP Constructor Promotion, Readonly Properties & Hooks — PHP 8.0 promotion, PHP 8.1 readonly, PHP 8.2 readonly class, PHP 8.4 hooks
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php80_constructor_property_promotion() {
    let out = run_prints(
        r#"<?php
class User {
    public function __construct(
        public string $name,
        public int $age = 18
    ) {}
}

$u = new User("Alice", 25);
echo "{$u->name} is {$u->age}";
"#,
    );
    assert_eq!(out, vec!["Alice is 25"]);
}

#[test]
fn test_php81_readonly_property_initialization() {
    let out = run_prints(
        r#"<?php
class Dto {
    public readonly string $uuid;
    public function __construct(string $uuid) {
        $this->uuid = $uuid;
    }
}

$d = new Dto("abc-123");
echo $d->uuid;
"#,
    );
    assert_eq!(out, vec!["abc-123"]);
}

#[test]
fn test_php82_readonly_class_declaration() {
    let out = run_prints(
        r##"<?php
readonly class ImmutablePoint {
    public function __construct(
        public float $x,
        public float $y
    ) {}
}

$p = new ImmutablePoint(3.5, 7.2);
echo "Point({$p->x}, {$p->y})";
"##,
    );
    assert_eq!(out, vec!["Point(3.5, 7.2)"]);
}

#[test]
fn test_php84_property_hooks_get_set_syntax() {
    let out = run_prints(
        r#"<?php
class Person {
    public string $first = "John";
    public string $last = "Doe";
    
    public string $fullName {
        get => "{$this->first} {$this->last}";
    }
}

$p = new Person();
echo $p->fullName;
"#,
    );
    assert_eq!(out, vec!["John Doe"]);
}

#[test]
fn test_php_constructor_promotion_with_attributes() {
    compile_ok(
        r#"<?php
#[Attribute]
class Validate {
    public function __construct(public string $rule) {}
}

class Product {
    public function __construct(
        #[Validate("min:1")]
        public string $title,
        #[Validate("gt:0")]
        public float $price
    ) {}
}

$p = new Product("Widget", 19.99);
echo $p->title;
"#,
    );
}

#[test]
fn test_php_promoted_properties_visibility_modifiers() {
    compile_ok(
        r#"<?php
class Service {
    public function __construct(
        private string $secretKey,
        protected string $endpoint,
        public int $timeout = 30
    ) {}
    
    public function getEndpoint(): string {
        return $this->endpoint;
    }
}

$s = new Service("key_123", "https://api.example.com");
echo $s->getEndpoint();
"#,
    );
}

#[test]
fn test_php_readonly_property_unassigned_error() {
    compile_ok(
        r#"<?php
class Document {
    public readonly string $title;
}

$doc = new Document();
"#,
    );
}

#[test]
fn test_php_asymmetric_property_visibility_php84() {
    compile_ok(
        r#"<?php
class Account {
    public private(set) float $balance = 0.0;

    public function deposit(float $amount): void {
        $this->balance += $amount;
    }
}

$acc = new Account();
$acc->deposit(100.0);
echo $acc->balance;
"#,
    );
}

#[test]
fn test_php_property_hooks_backing_field() {
    compile_ok(
        r#"<?php
class Temperature {
    public float $celsius {
        set {
            if ($value < -273.15) {
                throw new InvalidArgumentException("Below absolute zero");
            }
            $this->celsius = $value;
        }
    }
}

$t = new Temperature();
$t->celsius = 25.0;
echo $t->celsius;
"#,
    );
}

#[test]
fn test_php_constructor_promotion_default_expressions() {
    compile_ok(
        r#"<?php
class Config {
    public function __construct(
        public array $options = ["debug" => true],
        public string $env = "production"
    ) {}
}

$c = new Config();
echo $c->env;
"#,
    );
}
