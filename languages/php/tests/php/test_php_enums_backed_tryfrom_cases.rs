use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Enums Backed Cases & Operations — Pure enums, backed tryFrom(), cases iteration, enum methods, interface implementation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php81_backed_enum_tryfrom_returns_null_on_invalid() {
    let out = run_prints(
        r#"<?php
enum OrderState: string {
    case Pending = "pending";
    case Processing = "processing";
    case Completed = "completed";
}

$state1 = OrderState::tryFrom("completed");
$state2 = OrderState::tryFrom("invalid_state");

echo ($state1 !== null ? $state1->name : "NULL") . " | " . ($state2 === null ? "NULL" : "NOT_NULL");
"#,
    );
    assert_eq!(out, vec!["Completed | NULL"]);
}

#[test]
fn test_php81_enum_cases_iteration_map() {
    let out = run_prints(
        r#"<?php
enum Level: int {
    case Low = 10;
    case Medium = 20;
    case High = 30;
}

$names = array_map(fn($case) => $case->name . "=" . $case->value, Level::cases());
echo implode(", ", $names);
"#,
    );
    assert_eq!(out, vec!["Low=10, Medium=20, High=30"]);
}

#[test]
fn test_php81_enum_implementing_interface_with_method() {
    let out = run_prints(
        r#"<?php
interface Labelled {
    public function label(): string;
}

enum Currency: string implements Labelled {
    case USD = "USD";
    case EUR = "EUR";
    case GBP = "GBP";

    public function label(): string {
        return match($this) {
            self::USD => "US Dollar ($)",
            self::EUR => "Euro (€)",
            self::GBP => "Pound Sterling (£)",
        };
    }
}

echo Currency::EUR->label();
"#,
    );
    assert_eq!(out, vec!["Euro (€)"]);
}

#[test]
fn test_php81_pure_enum_comparison_identity() {
    let out = run_prints(
        r#"<?php
enum Action { case Create; case Update; case Delete; }

$a1 = Action::Create;
$a2 = Action::Create;
$a3 = Action::Update;

echo ($a1 === $a2 ? "1" : "0");
echo ($a1 === $a3 ? "1" : "0");
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_php81_enum_in_switch_case() {
    compile_ok(
        r#"<?php
enum Mode { case Read; case Write; case Admin; }

$m = Mode::Write;
$desc = "";
switch ($m) {
    case Mode::Read: $desc = "Read Only"; break;
    case Mode::Write: $desc = "Read Write"; break;
    case Mode::Admin: $desc = "Full Access"; break;
}
echo $desc;
"#,
    );
}

#[test]
fn test_php81_enum_constant_access() {
    compile_ok(
        r#"<?php
enum Feature: string {
    case Beta = "beta";
    case Stable = "stable";

    public const DEFAULT_FEATURE = self::Stable;
}

echo Feature::DEFAULT_FEATURE->value;
"#,
    );
}

#[test]
fn test_php81_enum_static_method_lookup() {
    compile_ok(
        r#"<?php
enum Severity: int {
    case Low = 1;
    case Medium = 2;
    case High = 3;

    public static function fromName(string $name): ?self {
        foreach (self::cases() as $case) {
            if ($case->name === $name) return $case;
        }
        return null;
    }
}

$s = Severity::fromName("High");
echo $s ? $s->value : 0;
"#,
    );
}

#[test]
fn test_php81_enum_json_encode_backed_value() {
    compile_ok(
        r#"<?php
enum Status: string { case Active = "active"; }
echo json_encode(Status::Active); // Enums serialize to backed value or name
"#,
    );
}

#[test]
fn test_php81_enum_reflection_backed_type() {
    compile_ok(
        r#"<?php
enum Code: int { case OK = 200; }
$re = new ReflectionEnum(Code::class);
$type = $re->getBackingType();
echo $type->getName();
"#,
    );
}

#[test]
fn test_php81_enum_in_type_hint_validation() {
    compile_ok(
        r#"<?php
enum Environment: string { case Dev = "dev"; case Prod = "prod"; }

function setEnv(Environment $env): string {
    return "Setting: " . $env->value;
}

echo setEnv(Environment::Dev);
"#,
    );
}
