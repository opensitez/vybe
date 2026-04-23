use super::helpers::compile_ok;

// ── Enum with traits ──────────────────────────────────────────

#[test] fn enum_with_trait() {
    compile_ok(r#"<?php
trait HasLabel {
    public function label(): string {
        return ucfirst(strtolower($this->name));
    }
}
enum Status {
    use HasLabel;
    case Active;
    case Inactive;
    case Pending;
}
echo Status::Active->label();
echo Status::Pending->label();
"#);
}

#[test] fn backed_enum_with_trait() {
    compile_ok(r#"<?php
trait Describable {
    public function describe(): string {
        return "{$this->name}={$this->value}";
    }
}
enum Color: string {
    use Describable;
    case Red   = 'red';
    case Green = 'green';
    case Blue  = 'blue';
}
echo Color::Red->describe();
"#);
}

#[test] fn enum_trait_with_property_check() {
    compile_ok(r#"<?php
trait HasPriority {
    public function isHighPriority(): bool {
        return match($this) {
            self::Critical, self::High => true,
            default => false,
        };
    }
}
enum Severity {
    use HasPriority;
    case Critical;
    case High;
    case Medium;
    case Low;
}
echo Severity::Critical->isHighPriority() ? 'high' : 'low';
echo Severity::Low->isHighPriority() ? 'high' : 'low';
"#);
}

// ── Enum implementing interface ───────────────────────────────

#[test] fn enum_implements_interface_complex() {
    compile_ok(r#"<?php
interface HasDisplayName {
    public function displayName(): string;
}
interface HasIcon {
    public function icon(): string;
}
enum FileType: string implements HasDisplayName, HasIcon {
    case PDF   = 'pdf';
    case Word  = 'docx';
    case Excel = 'xlsx';
    public function displayName(): string {
        return match($this) {
            self::PDF   => 'PDF Document',
            self::Word  => 'Word Document',
            self::Excel => 'Excel Spreadsheet',
        };
    }
    public function icon(): string {
        return match($this) {
            self::PDF   => '📄',
            self::Word  => '📝',
            self::Excel => '📊',
        };
    }
}
echo FileType::PDF->displayName();
echo FileType::Word->value;
"#);
}

// ── Enum as array keys ────────────────────────────────────────

#[test] fn enum_as_array_key_via_value() {
    compile_ok(r#"<?php
enum HttpMethod: string {
    case GET    = 'GET';
    case POST   = 'POST';
    case PUT    = 'PUT';
    case DELETE = 'DELETE';
}
$handlers = [
    HttpMethod::GET->value    => fn() => 'list',
    HttpMethod::POST->value   => fn() => 'create',
    HttpMethod::DELETE->value => fn() => 'delete',
];
$method = HttpMethod::POST;
echo ($handlers[$method->value])();
"#);
}

#[test] fn enum_in_match_exhaustive() {
    compile_ok(r#"<?php
enum Direction { case North; case South; case East; case West; }
function opposite(Direction $d): Direction {
    return match($d) {
        Direction::North => Direction::South,
        Direction::South => Direction::North,
        Direction::East  => Direction::West,
        Direction::West  => Direction::East,
    };
}
echo opposite(Direction::North)->name;
echo opposite(Direction::East)->name;
"#);
}

// ── Enum constants ────────────────────────────────────────────

#[test] fn enum_constants_basic() {
    compile_ok(r#"<?php
enum Suit: string {
    case Hearts   = 'H';
    case Diamonds = 'D';
    case Clubs    = 'C';
    case Spades   = 'S';
    const array RED_SUITS  = [self::Hearts, self::Diamonds];
    const array BLACK_SUITS = [self::Clubs, self::Spades];
}
echo count(Suit::RED_SUITS) . ':' . count(Suit::BLACK_SUITS);
"#);
}

#[test] fn enum_constant_expressions() {
    compile_ok(r#"<?php
enum Permission: int {
    case Read    = 1;
    case Write   = 2;
    case Execute = 4;
    const int ALL = self::Read->value | self::Write->value | self::Execute->value;
}
echo Permission::ALL;
"#);
}

// ── cases() filtering ─────────────────────────────────────────

#[test] fn enum_cases_filter() {
    compile_ok(r#"<?php
enum Priority: int {
    case Low    = 1;
    case Medium = 2;
    case High   = 3;
    case Critical = 4;
    public function isUrgent(): bool { return $this->value >= 3; }
}
$urgent = array_filter(Priority::cases(), fn($c) => $c->isUrgent());
echo count($urgent);
echo ':' . implode(',', array_map(fn($c) => $c->name, $urgent));
"#);
}

#[test] fn enum_cases_map() {
    compile_ok(r#"<?php
enum Weekday: int {
    case Monday    = 1;
    case Tuesday   = 2;
    case Wednesday = 3;
    case Thursday  = 4;
    case Friday    = 5;
    case Saturday  = 6;
    case Sunday    = 7;
    public function isWeekend(): bool { return $this->value >= 6; }
}
$weekends = array_filter(Weekday::cases(), fn($d) => $d->isWeekend());
$names = array_map(fn($d) => $d->name, $weekends);
echo implode(',', $names);
"#);
}

// ── Enum methods complex ──────────────────────────────────────

#[test] fn enum_method_next_prev() {
    compile_ok(r#"<?php
enum Month: int {
    case January  = 1; case February = 2; case March    = 3;
    case April    = 4; case May      = 5; case June     = 6;
    case July     = 7; case August   = 8; case September = 9;
    case October  = 10; case November = 11; case December = 12;
    public function next(): self {
        $next = ($this->value % 12) + 1;
        return self::from($next);
    }
    public function daysInMonth(int $year = 2024): int {
        return cal_days_in_month(CAL_GREGORIAN, $this->value, $year);
    }
}
echo Month::December->next()->name;
echo Month::February->daysInMonth(2024);
"#);
}

#[test] fn enum_method_comparison() {
    compile_ok(r#"<?php
enum Size: int {
    case XS = 1; case S = 2; case M = 3; case L = 4; case XL = 5;
    public function fitsInto(self $other): bool { return $this->value <= $other->value; }
    public function between(self $min, self $max): bool {
        return $this->value >= $min->value && $this->value <= $max->value;
    }
}
echo Size::S->fitsInto(Size::L) ? 'fits' : 'no fit';
echo Size::M->between(Size::S, Size::XL) ? ':in range' : ':out of range';
"#);
}

// ── from / tryFrom deep ───────────────────────────────────────

#[test] fn backed_enum_from_valid() {
    compile_ok(r#"<?php
enum Status: string {
    case Active   = 'active';
    case Inactive = 'inactive';
    case Banned   = 'banned';
}
$s = Status::from('active');
echo $s->name . ':' . $s->value;
"#);
}

#[test] fn backed_enum_try_from_invalid() {
    compile_ok(r#"<?php
enum Code: int { case OK = 200; case NotFound = 404; case Error = 500; }
$found  = Code::tryFrom(200);
$missing = Code::tryFrom(999);
echo ($found !== null ? $found->name : 'null') . ':';
echo ($missing !== null ? $missing->name : 'null');
"#);
}

#[test] fn backed_enum_from_user_input() {
    compile_ok(r#"<?php
enum Language: string {
    case PHP    = 'php';
    case Python = 'python';
    case Rust   = 'rust';
}
$inputs = ['php', 'rust', 'javascript', 'python'];
foreach ($inputs as $input) {
    $lang = Language::tryFrom($input);
    echo ($lang !== null ? $lang->name : 'unknown') . ' ';
}
"#);
}

// ── Enum in collections ───────────────────────────────────────

#[test] fn enum_in_array_collect() {
    compile_ok(r#"<?php
enum Fruit { case Apple; case Banana; case Cherry; }
$basket = [Fruit::Apple, Fruit::Banana, Fruit::Apple, Fruit::Cherry];
$counts = [];
foreach ($basket as $fruit) {
    $counts[$fruit->name] = ($counts[$fruit->name] ?? 0) + 1;
}
ksort($counts);
foreach ($counts as $name => $count) { echo "$name:$count "; }
"#);
}

#[test] fn enum_sorted_by_value() {
    compile_ok(r#"<?php
enum Priority: int { case Low = 1; case Medium = 5; case High = 10; }
$tasks = [
    ['name' => 'cleanup', 'priority' => Priority::Low],
    ['name' => 'deploy',  'priority' => Priority::High],
    ['name' => 'review',  'priority' => Priority::Medium],
];
usort($tasks, fn($a, $b) => $b['priority']->value <=> $a['priority']->value);
foreach ($tasks as $task) { echo $task['name'] . ' '; }
"#);
}

// ── Enum name and value access ────────────────────────────────

#[test] fn enum_name_value_both() {
    compile_ok(r#"<?php
enum Currency: string {
    case USD = 'US Dollar';
    case EUR = 'Euro';
    case GBP = 'British Pound';
}
foreach (Currency::cases() as $c) {
    echo "{$c->name}: {$c->value}\n";
}
"#);
}

#[test] fn pure_enum_name_only() {
    compile_ok(r#"<?php
enum Planet { case Mercury; case Venus; case Earth; case Mars; }
$names = array_map(fn($p) => $p->name, Planet::cases());
echo implode(',', $names);
"#);
}

// ── Enum in static context ────────────────────────────────────

#[test] fn enum_static_method_factory() {
    compile_ok(r#"<?php
enum Environment: string {
    case Development = 'dev';
    case Staging     = 'staging';
    case Production  = 'prod';
    public static function fromEnvVar(): self {
        return self::tryFrom(getenv('APP_ENV') ?: '') ?? self::Development;
    }
    public function isProduction(): bool { return $this === self::Production; }
}
$env = Environment::fromEnvVar();
echo $env->name;
echo $env->isProduction() ? ':prod' : ':not prod';
"#);
}
