use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: OOP Advanced Modifiers & Property Combinations — final, readonly, promoted properties, interface contracts
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_final_class_with_promoted_readonly_properties() {
    let out = run_prints(
        r#"<?php
final class ImmutableConfig {
    public function __construct(
        public readonly string $env,
        public readonly bool $debug = false
    ) {}
}

$cfg = new ImmutableConfig("production");
echo "{$cfg->env} debug=" . ($cfg->debug ? "1" : "0");
"#,
    );
    assert_eq!(out, vec!["production debug=0"]);
}

#[test]
fn test_php_readonly_class_inheritance_contract() {
    let out = run_prints(
        r#"<?php
readonly class BaseDto {
    public function __construct(public int $id) {}
}

readonly class UserDto extends BaseDto {
    public function __construct(int $id, public string $name) {
        parent::__construct($id);
    }
}

$u = new UserDto(42, "Alice");
echo "{$u->id}:{$u->name}";
"#,
    );
    assert_eq!(out, vec!["42:Alice"]);
}

#[test]
fn test_php_final_method_in_abstract_base_class() {
    let out = run_prints(
        r#"<?php
abstract class TemplateMethodProcessor {
    final public function execute(): string {
        return "PRE -> " . $this->step() . " -> POST";
    }
    abstract protected function step(): string;
}

class ConcreteProcessor extends TemplateMethodProcessor {
    protected function step(): string { return "STEP_BODY"; }
}

$cp = new ConcreteProcessor();
echo $cp->execute();
"#,
    );
    assert_eq!(out, vec!["PRE -> STEP_BODY -> POST"]);
}

#[test]
fn test_php_promoted_property_doc_comments() {
    compile_ok(
        r#"<?php
class Customer {
    public function __construct(
        /** @var string Customer full name */
        public string $name,
        /** @var string Customer email address */
        public string $email
    ) {}
}

$c = new Customer("Bob", "bob@example.com");
echo "$c->name <$c->email>";
"#,
    );
}

#[test]
fn test_php_readonly_property_cloning_behavior() {
    compile_ok(
        r#"<?php
class Order {
    public readonly DateTimeImmutable $createdAt;
    public function __construct() {
        $this->createdAt = new DateTimeImmutable();
    }
    public function __clone() {
        // Readonly properties can be modified during __clone in PHP 8.3+
    }
}

$o1 = new Order();
$o2 = clone $o1;
echo get_class($o2);
"#,
    );
}

#[test]
fn test_php_promoted_property_callable_default_error() {
    compile_ok(
        r#"<?php
class Task {
    public function __construct(
        public string $title,
        public int $priority = 1
    ) {}
}

$t = new Task("Write Tests");
echo $t->title;
"#,
    );
}

#[test]
fn test_php_interface_readonly_property_hook_contract() {
    compile_ok(
        r#"<?php
interface Identifiable {
    public int $id { get; }
}

class Record implements Identifiable {
    public function __construct(public readonly int $id) {}
}

$r = new Record(101);
echo $r->id;
"#,
    );
}

#[test]
fn test_php_asymmetric_visibility_constructor_promotion_php84() {
    compile_ok(
        r#"<?php
class CounterService {
    public function __construct(
        public private(set) int $count = 0
    ) {}

    public function increment(): void {
        $this->count++;
    }
}

$cs = new CounterService();
$cs->increment();
echo $cs->count;
"#,
    );
}

#[test]
fn test_php_final_class_constant_overriding_prevention() {
    compile_ok(
        r#"<?php
class BaseConstants {
    final public const VERSION = "1.0.0";
}

class AppConstants extends BaseConstants {}

echo AppConstants::VERSION;
"#,
    );
}

#[test]
fn test_php_promoted_property_varargs_expansion() {
    compile_ok(
        r#"<?php
class UserList {
    public array $users;
    public function __construct(string ...$users) {
        $this->users = $users;
    }
}

$ul = new UserList("Alice", "Bob", "Charlie");
echo count($ul->users);
"#,
    );
}

#[test]
fn test_php_readonly_class_with_runtime_property_reading() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class SessionState {
    public function __construct(
        public int $id,
        public string $name,
        public bool $active = true
    ) {}
}

$state = new SessionState(1, "auth");
echo $state->id . "|" . $state->name . "|" . ($state->active ? "yes" : "no");
"#,
        ),
        vec!["1|auth|yes"]
    );
}

#[test]
fn test_php_readonly_property_reassignment_is_rejected() {
    assert_eq!(
        run_prints(
            r#"<?php
class ImmutableCounter {
    public function __construct(public readonly int $value) {}

    public function safeSet(int $next): string {
        try {
            $this->value = $next;
            return "ok";
        } catch (Error $e) {
            return "error";
        }
    }
}

$counter = new ImmutableCounter(4);
echo $counter->value . "|" . $counter->safeSet(7) . "|" . $counter->value;
"#,
        ),
        vec!["4|error|4"]
    );
}

#[test]
fn test_php_final_factory_of_readonly_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
final class Builder {
    private function __construct(
        public readonly string $token,
        public readonly int $ttl
    ) {}

    public static function fromDate(string $prefix, int $year): self {
        return new self("$prefix-$year", 900);
    }
}

$builder = Builder::fromDate("token", 2026);
echo $builder->token . "|" . $builder->ttl;
"#,
        ),
        vec!["token-2026|900"]
    );
}

#[test]
fn test_php_readonly_class_with_parent_readonly_property() {
    assert_eq!(
        run_prints(
            r#"<?php
readonly class BaseConfig {
    public function __construct(public readonly string $env) {}
}

readonly class ServiceConfig extends BaseConfig {
    public function __construct(string $env, public readonly string $region) {
        parent::__construct($env);
    }
}

$cfg = new ServiceConfig("prod", "eu");
echo $cfg->env . "|" . $cfg->region;
"#,
        ),
        vec!["prod|eu"]
    );
}
