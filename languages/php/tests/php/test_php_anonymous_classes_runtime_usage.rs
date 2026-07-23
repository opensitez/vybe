use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Anonymous Classes Runtime & Mechanics — new class($arg) extends Base implements Interface, constructor, attributes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_anonymous_class_instantiation_with_constructor() {
    let out = run_prints(
        r#"<?php
$logger = new class("APP_LOG") {
    public function __construct(public string $prefix) {}
    public function info(string $msg): string {
        return "[{$this->prefix}] $msg";
    }
};

echo $logger->info("Service started");
"#,
    );
    assert_eq!(out, vec!["[APP_LOG] Service started"]);
}

#[test]
fn test_php_anonymous_class_extending_base_class() {
    let out = run_prints(
        r#"<?php
abstract class BaseHandler {
    abstract public function handle(): string;
}

$handler = new class extends BaseHandler {
    public function handle(): string {
        return "Handled by anonymous subclass";
    }
};

echo $handler->handle();
"#,
    );
    assert_eq!(out, vec!["Handled by anonymous subclass"]);
}

#[test]
fn test_php_anonymous_class_implementing_interface() {
    let out = run_prints(
        r#"<?php
interface Renderable {
    public function render(): string;
}

$component = new class implements Renderable {
    public function render(): string {
        return "<span>HTML</span>";
    }
};

echo ($component instanceof Renderable ? "IS_RENDERABLE: " : "NO: ") . $component->render();
"#,
    );
    assert_eq!(out, vec!["IS_RENDERABLE: <span>HTML</span>"]);
}

#[test]
fn test_php_anonymous_class_outer_scope_capture_via_constructor() {
    let out = run_prints(
        r#"<?php
$config = ["debug" => true, "env" => "staging"];

$service = new class($config) {
    public function __construct(private array $cfg) {}
    public function getEnv(): string { return $this->cfg["env"]; }
};

echo $service->getEnv();
"#,
    );
    assert_eq!(out, vec!["staging"]);
}

#[test]
fn test_php_anonymous_class_nested_inside_class_method() {
    compile_ok(
        r#"<?php
class Container {
    public function createStrategy(): object {
        return new class {
            public function execute(): string { return "Strategy executed"; }
        };
    }
}

$c = new Container();
$strat = $c->createStrategy();
echo $strat->execute();
"#,
    );
}

#[test]
fn test_php_anonymous_class_using_traits() {
    compile_ok(
        r#"<?php
trait IdentityTrait {
    public function getId(): int { return 999; }
}

$entity = new class {
    use IdentityTrait;
};

echo $entity->getId();
"#,
    );
}

#[test]
fn test_php_anonymous_class_name_format_inspection() {
    compile_ok(
        r#"<?php
$anon = new class {};
$className = get_class($anon);
echo str_contains($className, "class@anonymous") ? "ANONYMOUS_NAME_OK" : "NAMED";
"#,
    );
}

#[test]
fn test_php_anonymous_class_readonly_php82() {
    compile_ok(
        r#"<?php
$immutable = new readonly class(10, 20) {
    public function __construct(public int $x, public int $y) {}
};

echo "{$immutable->x}, {$immutable->y}";
"#,
    );
}

#[test]
fn test_php_anonymous_class_with_attribute_annotation() {
    compile_ok(
        r#"<?php
#[Attribute]
class ServiceMeta { public function __construct(public string $type) {} }

$service = new #[ServiceMeta("transient")] class {
    public function run() { return "ok"; }
};

$rc = new ReflectionClass($service);
echo count($rc->getAttributes(ServiceMeta::class));
"#,
    );
}

#[test]
fn test_php_anonymous_class_in_return_type_hint() {
    compile_ok(
        r#"<?php
function createAnonymous(): object {
    return new class {
        public string $status = "active";
    };
}

$obj = createAnonymous();
echo $obj->status;
"#,
    );
}
