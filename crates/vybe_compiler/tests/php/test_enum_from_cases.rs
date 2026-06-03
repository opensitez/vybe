use super::helpers::{compile_ok, run_prints};

// ── Backed enum from() ────────────────────────────────────────

#[test]
fn backed_enum_from_valid_int() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status: int {
    case Active = 1;
    case Inactive = 0;
}
$s = Status::from(1);
echo $s->name;
"#
        ),
        vec!["Active"]
    );
}

#[test]
fn backed_enum_from_valid_string() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Color: string {
    case Red = 'red';
    case Blue = 'blue';
}
$c = Color::from('blue');
echo $c->value;
"#
        ),
        vec!["blue"]
    );
}

#[test]
fn backed_enum_from_throws_on_invalid() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status: int { case Active = 1; }
try {
    Status::from(99);
} catch (\ValueError $e) {
    echo "error";
}
"#
        ),
        vec!["error"]
    );
}

// ── Backed enum tryFrom() ─────────────────────────────────────

#[test]
fn backed_enum_try_from_returns_case_on_match() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Priority: int { case Low = 1; case High = 3; }
$p = Priority::tryFrom(3);
echo $p->name;
"#
        ),
        vec!["High"]
    );
}

#[test]
fn backed_enum_try_from_returns_null_on_no_match() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Priority: int { case Low = 1; case High = 3; }
$p = Priority::tryFrom(99);
echo var_export($p, true);
"#
        ),
        vec!["NULL"]
    );
}

#[test]
fn backed_enum_try_from_string_no_match() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Suit: string { case Hearts = 'H'; case Clubs = 'C'; }
$s = Suit::tryFrom('X');
echo ($s === null) ? 'null' : 'found';
"#
        ),
        vec!["null"]
    );
}

// ── Enum::cases() ─────────────────────────────────────────────

#[test]
fn pure_enum_cases_returns_all() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Direction { case North; case South; case East; case West; }
$cases = Direction::cases();
echo count($cases);
"#
        ),
        vec!["4"]
    );
}

#[test]
fn backed_enum_cases_returns_case_objects() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status: string { case Active = 'A'; case Inactive = 'I'; }
$names = array_map(fn($c) => $c->name, Status::cases());
sort($names);
echo implode(',', $names);
"#
        ),
        vec!["Active,Inactive"]
    );
}

#[test]
fn enum_cases_preserves_declaration_order() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Step { case First; case Second; case Third; }
$names = array_map(fn($c) => $c->name, Step::cases());
echo implode(',', $names);
"#
        ),
        vec!["First,Second,Third"]
    );
}

// ── Enum implements interface ─────────────────────────────────

#[test]
fn enum_implements_interface_with_method() {
    assert_eq!(
        run_prints(
            r#"<?php
interface HasLabel {
    public function label(): string;
}
enum Color: string implements HasLabel {
    case Red = 'red';
    case Blue = 'blue';
    public function label(): string { return ucfirst($this->value); }
}
echo Color::Red->label();
"#
        ),
        vec!["Red"]
    );
}

#[test]
fn enum_implements_stringable() {
    assert_eq!(
        run_prints(
            r#"<?php
interface HasLabel {
    public function label(): string;
}
enum Suit: string implements HasLabel {
    case Hearts = 'hearts';
    public function label(): string { return $this->name . '(' . $this->value . ')'; }
}
echo Suit::Hearts->label();
"#
        ),
        vec!["Hearts(hearts)"]
    );
}

// ── Enum methods ──────────────────────────────────────────────

#[test]
fn enum_method_access_value() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Priority: int {
    case Low = 1;
    case Medium = 5;
    case High = 10;
    public function isUrgent(): bool { return $this->value >= 10; }
}
echo Priority::High->isUrgent() ? 'urgent' : 'normal';
"#
        ),
        vec!["urgent"]
    );
}

#[test]
fn enum_method_returns_other_case() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Toggle { case On; case Off;
    public function flip(): self {
        return match($this) { Toggle::On => Toggle::Off, Toggle::Off => Toggle::On };
    }
}
echo Toggle::On->flip()->name;
"#
        ),
        vec!["Off"]
    );
}

#[test]
fn enum_static_method() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Level: int {
    case Low = 1; case High = 2;
    public static function default(): self { return self::Low; }
}
echo Level::default()->name;
"#
        ),
        vec!["Low"]
    );
}

// ── Enum constants ────────────────────────────────────────────

#[test]
fn enum_can_have_constants() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Suit { case Hearts; case Clubs;
    const DEFAULT = self::Hearts;
}
echo Suit::DEFAULT->name;
"#
        ),
        vec!["Hearts"]
    );
}

// ── Enum in match expression ──────────────────────────────────

#[test]
fn enum_in_match_arm() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Season { case Spring; case Summer; case Autumn; case Winter; }
$s = Season::Autumn;
echo match($s) {
    Season::Spring, Season::Summer => 'warm',
    Season::Autumn, Season::Winter => 'cold',
};
"#
        ),
        vec!["cold"]
    );
}

// ── Enum comparison ───────────────────────────────────────────

#[test]
fn enum_cases_are_identical_singletons() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Color { case Red; case Blue; }
echo (Color::Red === Color::Red) ? 'same' : 'diff';
"#
        ),
        vec!["same"]
    );
}

#[test]
fn enum_different_cases_not_identical() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Color { case Red; case Blue; }
echo (Color::Red === Color::Blue) ? 'same' : 'diff';
"#
        ),
        vec!["diff"]
    );
}

#[test]
fn backed_enum_value_accessible() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Code: int { case OK = 200; case NotFound = 404; }
echo Code::NotFound->value;
"#
        ),
        vec!["404"]
    );
}

// ── Enum in array ─────────────────────────────────────────────

#[test]
fn enum_cases_in_array_with_in_array() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Fruit { case Apple; case Banana; case Cherry; }
$allowed = [Fruit::Apple, Fruit::Cherry];
echo in_array(Fruit::Banana, $allowed) ? 'yes' : 'no';
"#
        ),
        vec!["no"]
    );
}

// ── Enum with trait ───────────────────────────────────────────

#[test]
fn enum_uses_trait() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Describable {
    public function describe(): string { return "I am " . $this->name; }
}
enum Planet { case Mars; case Venus; use Describable; }
echo Planet::Mars->describe();
"#
        ),
        vec!["I am Mars"]
    );
}

// ── Backed enum in function parameter ────────────────────────

#[test]
fn backed_enum_as_type_hint_parameter() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status: string { case Active = 'active'; case Banned = 'banned'; }
function getStatusLabel(Status $s): string {
    return match($s) { Status::Active => 'Active User', Status::Banned => 'Banned User' };
}
echo getStatusLabel(Status::Active);
"#
        ),
        vec!["Active User"]
    );
}

// ── Enum name property ────────────────────────────────────────

#[test]
fn pure_enum_name_property() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Direction { case North; case South; }
echo Direction::North->name;
"#
        ),
        vec!["North"]
    );
}

#[test]
fn backed_enum_name_and_value_distinct() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Http: int { case Ok = 200; }
$c = Http::Ok;
echo $c->name . ':' . $c->value;
"#
        ),
        vec!["Ok:200"]
    );
}
