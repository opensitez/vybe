use super::helpers::run_prints;

// ── Enum methods ──────────────────────────────────────────────

#[test]
fn enum_method_on_case() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Direction {
    case North; case South; case East; case West;
    public function opposite(): self {
        return match($this) {
            self::North => self::South,
            self::South => self::North,
            self::East => self::West,
            self::West => self::East,
        };
    }
}
echo Direction::North->opposite()->name;
"#
        ),
        vec!["South"]
    );
}
#[test]
fn enum_static_method() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Color: string {
    case Red = 'red'; case Green = 'green'; case Blue = 'blue';
    public static function fromHex(string $hex): self {
        return match($hex) { '#ff0000' => self::Red, '#00ff00' => self::Green, default => self::Blue };
    }
}
echo Color::fromHex('#ff0000')->value;
"#
        ),
        vec!["red"]
    );
}
#[test]
fn enum_method_with_logic() {
    assert_eq!(
        run_prints(
            r#"<?php
enum HttpMethod: string {
    case GET = 'GET'; case POST = 'POST'; case PUT = 'PUT'; case DELETE = 'DELETE';
    public function isSafe(): bool { return match($this) { self::GET => true, default => false }; }
    public function isIdempotent(): bool { return match($this) { self::GET, self::PUT, self::DELETE => true, default => false }; }
}
echo HttpMethod::GET->isSafe() ? 'safe' : 'unsafe';
echo ',' . (HttpMethod::POST->isIdempotent() ? 'idem' : 'not');
"#
        ),
        vec!["safe,not"]
    );
}

// ── Enum interfaces ───────────────────────────────────────────

#[test]
fn enum_implements_interface_method() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Colorable { public function hex(): string; }
enum Palette: string implements Colorable {
    case Red = 'red'; case Blue = 'blue';
    public function hex(): string { return match($this) { self::Red => '#FF0000', self::Blue => '#0000FF' }; }
}
echo Palette::Red->hex();
"#
        ),
        vec!["#FF0000"]
    );
}

// ── Enum cases() ──────────────────────────────────────────────

#[test]
fn enum_cases_returns_all() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Day { case Mon; case Tue; case Wed; case Thu; case Fri; case Sat; case Sun; }
$cases = Day::cases();
echo count($cases) . ':' . $cases[0]->name;
"#
        ),
        vec!["7:Mon"]
    );
}
#[test]
fn enum_backed_cases_values() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Priority: int { case Low = 1; case Mid = 2; case High = 3; }
echo implode(',', array_map(fn($c) => $c->value, Priority::cases()));
"#
        ),
        vec!["1,2,3"]
    );
}

// ── Enum in array / data structures ──────────────────────────

#[test]
fn enum_as_array_key() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Fruit { case Apple; case Banana; case Cherry; }
$prices = [Fruit::Apple->name => 1.5, Fruit::Banana->name => 0.5];
echo $prices[Fruit::Apple->name];
"#
        ),
        vec!["1.5"]
    );
}
#[test]
fn enum_in_match_multiple_arms() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Suit { case Hearts; case Diamonds; case Clubs; case Spades; }
function color(Suit $s): string {
    return match($s) { Suit::Hearts, Suit::Diamonds => 'red', Suit::Clubs, Suit::Spades => 'black' };
}
echo color(Suit::Diamonds) . ',' . color(Suit::Clubs);
"#
        ),
        vec!["red,black"]
    );
}
#[test]
fn enum_collection_filter() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status: int { case Active = 1; case Inactive = 0; case Pending = 2; }
$active = array_filter(Status::cases(), fn($c) => $c->value > 0);
echo count($active);
"#
        ),
        vec!["2"]
    );
}

// ── Enum constants ────────────────────────────────────────────

#[test]
fn enum_with_constants() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Planet {
    case Mercury; case Venus; case Earth;
    const HABITABLE = self::Earth;
}
echo Planet::HABITABLE->name;
"#
        ),
        vec!["Earth"]
    );
}
#[test]
#[allow(non_snake_case)]
fn backed_enum_from_tryFrom() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Weekday: int { case Mon=1; case Tue=2; case Wed=3; case Thu=4; case Fri=5; case Sat=6; case Sun=7; }
echo Weekday::from(3)->name;
echo ',' . (Weekday::tryFrom(99)?->name ?? 'none');
"#
        ),
        vec!["Wed,none"]
    );
}

// ── Enum comparison ───────────────────────────────────────────

#[test]
fn enum_singletons_compared_by_identity() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status { case Active; case Inactive; }
$a = Status::Active;
$b = Status::Active;
echo ($a === $b) ? 'same' : 'diff';
"#
        ),
        vec!["same"]
    );
}
#[test]
fn enum_different_cases_not_equal() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Coin { case Heads; case Tails; }
echo (Coin::Heads === Coin::Tails) ? 'eq' : 'neq';
"#
        ),
        vec!["neq"]
    );
}

// ── Enum in type declarations ─────────────────────────────────

#[test]
fn enum_as_parameter_type() {
    assert_eq!(
        run_prints(
            r#"<?php
enum LogLevel { case Debug; case Info; case Warning; case Error; }
function log2(LogLevel $level, string $msg): void {
    echo "[{$level->name}] $msg";
}
log2(LogLevel::Error, 'crash');
"#
        ),
        vec!["[Error] crash"]
    );
}
#[test]
fn enum_nullable_parameter() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Mode { case Fast; case Safe; }
function run(?Mode $m): string { return $m === null ? 'default' : $m->name; }
echo run(null) . ',' . run(Mode::Fast);
"#
        ),
        vec!["default,Fast"]
    );
}
