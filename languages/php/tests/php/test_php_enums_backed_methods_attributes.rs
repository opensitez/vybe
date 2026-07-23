use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Enums & Attributes — Pure enums, backed enums, from(), tryFrom(), methods in enums, #[Attribute]
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php81_string_backed_enum_from_tryfrom() {
    let out = run_prints(
        r#"<?php
enum Status: string {
    case Pending = "pending";
    case Active = "active";
}

$s = Status::from("active");
echo $s->name . "=" . $s->value;
"#,
    );
    assert_eq!(out, vec!["Active=active"]);
}

#[test]
fn test_php81_int_backed_enum_cases() {
    let out = run_prints(
        r#"<?php
enum HTTPStatus: int {
    case OK = 200;
    case NotFound = 404;
}

$status = HTTPStatus::tryFrom(404);
echo $status ? $status->name : "NULL";
"#,
    );
    assert_eq!(out, vec!["NotFound"]);
}

#[test]
fn test_php81_enum_custom_methods_and_interfaces() {
    let out = run_prints(
        r#"<?php
interface Colorable {
    public function color(): string;
}

enum Priority: int implements Colorable {
    case Low = 1;
    case Medium = 2;
    case High = 3;

    public function color(): string {
        return match($this) {
            self::Low => "green",
            self::Medium => "yellow",
            self::High => "red",
        };
    }
}

echo Priority::High->color();
"#,
    );
    assert_eq!(out, vec!["red"]);
}

#[test]
fn test_php80_attribute_declaration_instantiation() {
    let out = run_prints(
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS | Attribute::TARGET_METHOD)]
class Route {
    public function __construct(
        public string $path,
        public string $method = "GET"
    ) {}
}

#[Route("/api/users", method: "POST")]
class UserController {}

$rc = new ReflectionClass(UserController::class);
$attrs = $rc->getAttributes(Route::class);
$route = $attrs[0]->newInstance();

echo "{$route->method} {$route->path}";
"#,
    );
    assert_eq!(out, vec!["POST /api/users"]);
}

#[test]
fn test_php81_enum_cases_array() {
    compile_ok(
        r#"<?php
enum Direction {
    case North;
    case South;
    case East;
    case West;
}

$cases = Direction::cases();
foreach ($cases as $case) {
    echo $case->name . "\n";
}
"#,
    );
}

#[test]
fn test_php81_enum_static_methods() {
    compile_ok(
        r#"<?php
enum Role: string {
    case Admin = "admin";
    case User = "user";

    public static function values(): array {
        return array_column(self::cases(), "value");
    }
}

print_r(Role::values());
"#,
    );
}

#[test]
fn test_php81_repeatable_attribute() {
    compile_ok(
        r#"<?php
#[Attribute(Attribute::TARGET_METHOD | Attribute::IS_REPEATABLE)]
class Middleware {
    public function __construct(public string $name) {}
}

class DashboardController {
    #[Middleware("auth")]
    #[Middleware("log")]
    public function index() {}
}

$rm = new ReflectionMethod(DashboardController::class, "index");
$attrs = $rm->getAttributes(Middleware::class);
echo count($attrs);
"#,
    );
}

#[test]
fn test_php81_nested_attributes_in_parameters() {
    compile_ok(
        r#"<?php
#[Attribute(Attribute::TARGET_PARAMETER)]
class Inject {
    public function __construct(public string $service) {}
}

class PaymentProcessor {
    public function __construct(
        #[Inject("db.connection")]
        public object $db
    ) {}
}

$rp = new ReflectionParameter([PaymentProcessor::class, "__construct"], "db");
echo count($rp->getAttributes(Inject::class));
"#,
    );
}

#[test]
fn test_php81_enum_in_match_expression() {
    compile_ok(
        r#"<?php
enum State { case Draft; case Published; case Archived; }

$state = State::Published;
$label = match($state) {
    State::Draft => "Draft Document",
    State::Published => "Live Document",
    State::Archived => "Archived",
};
echo $label;
"#,
    );
}

#[test]
fn test_php81_enum_constant_expressions() {
    compile_ok(
        r#"<?php
enum Size {
    case Small;
    case Medium;
    case Large;

    public const Default = self::Medium;
}

echo Size::Default->name;
"#,
    );
}
