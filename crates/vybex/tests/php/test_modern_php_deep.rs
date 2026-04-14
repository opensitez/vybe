use super::helpers::run_prints;

// ── Match expression (deep) ──────────────────────────────────────
#[test]
fn match_with_complex_arms() {
    assert_eq!(run_prints(r#"<?php
function classify(int $score): string {
    return match(true) {
        $score >= 90 => "A",
        $score >= 80 => "B",
        $score >= 70 => "C",
        $score >= 60 => "D",
        default => "F",
    };
}
echo classify(95);
echo classify(82);
echo classify(55);
"#), &["A", "B", "F"]);
}

#[test]
fn match_multiple_conditions() {
    assert_eq!(run_prints(r#"<?php
function httpStatus(int $code): string {
    return match($code) {
        200, 201 => "success",
        301, 302 => "redirect",
        404 => "not found",
        500, 502, 503 => "server error",
        default => "unknown",
    };
}
echo httpStatus(200);
echo httpStatus(301);
echo httpStatus(404);
echo httpStatus(503);
echo httpStatus(418);
"#), &["success", "redirect", "not found", "server error", "unknown"]);
}

#[test]
fn match_no_default_throws() {
    assert_eq!(run_prints(r#"<?php
try {
    $x = 5;
    $result = match($x) {
        1 => "one",
        2 => "two",
    };
} catch (\UnhandledMatchError $e) {
    echo "unhandled";
}
"#), &["unhandled"]);
}

// ── Fiber patterns (deep) ────────────────────────────────────────
#[test]
fn fiber_scheduler_round_robin() {
    assert_eq!(run_prints(r#"<?php
$fibers = [];
for ($i = 1; $i <= 3; $i++) {
    $n = $i;
    $fibers[] = new Fiber(function() use ($n) {
        echo "task$n start";
        Fiber::suspend();
        echo "task$n end";
    });
}
// Start all
foreach ($fibers as $f) {
    $f->start();
}
// Resume all
foreach ($fibers as $f) {
    $f->resume();
}
"#), &["task1 start", "task2 start", "task3 start", "task1 end", "task2 end", "task3 end"]);
}

#[test]
fn fiber_with_return() {
    assert_eq!(run_prints(r#"<?php
$fiber = new Fiber(function(): string {
    $a = Fiber::suspend("first");
    $b = Fiber::suspend("second");
    return "result: $a + $b";
});
echo $fiber->start();
echo $fiber->resume("hello");
echo $fiber->resume("world");
echo $fiber->getReturn();
"#), &["first", "second", "result: hello + world"]);
}

// ── First-class callables (deep) ─────────────────────────────────
#[test]
fn first_class_callable_strlen() {
    assert_eq!(run_prints(r#"<?php
$fn = strlen(...);
echo $fn("hello");
echo $fn("hi");
"#), &["5", "2"]);
}

#[test]
fn first_class_callable_method() {
    assert_eq!(run_prints(r#"<?php
class Formatter {
    public function upper(string $s): string {
        return strtoupper($s);
    }
}
$f = new Formatter();
$fn = $f->upper(...);
echo $fn("hello");
"#), &["HELLO"]);
}

#[test]
fn first_class_callable_in_map() {
    assert_eq!(run_prints(r#"<?php
$words = ["hello", "world", "php"];
$result = array_map(strtoupper(...), $words);
echo implode(",", $result);
"#), &["HELLO,WORLD,PHP"]);
}

// ── Enum advanced ────────────────────────────────────────────────
#[test]
fn enum_cases_list() {
    assert_eq!(run_prints(r#"<?php
enum Direction {
    case North;
    case South;
    case East;
    case West;
}
$cases = Direction::cases();
echo count($cases);
echo $cases[0]->name;
echo $cases[3]->name;
"#), &["4", "North", "West"]);
}

#[test]
fn enum_in_match() {
    assert_eq!(run_prints(r#"<?php
enum Suit: string {
    case Hearts = "H";
    case Diamonds = "D";
    case Clubs = "C";
    case Spades = "S";
}
function color(Suit $s): string {
    return match($s) {
        Suit::Hearts, Suit::Diamonds => "red",
        Suit::Clubs, Suit::Spades => "black",
    };
}
echo color(Suit::Hearts);
echo color(Suit::Spades);
"#), &["red", "black"]);
}

#[test]
fn enum_with_method_and_interface() {
    assert_eq!(run_prints(r#"<?php
interface HasSymbol {
    public function symbol(): string;
}
enum Currency: string implements HasSymbol {
    case USD = "usd";
    case EUR = "eur";
    case GBP = "gbp";
    public function symbol(): string {
        return match($this) {
            self::USD => "$",
            self::EUR => "€",
            self::GBP => "£",
        };
    }
}
echo Currency::USD->symbol();
echo Currency::EUR->symbol();
echo Currency::GBP->value;
"#), &["$", "€", "gbp"]);
}

// ── Readonly class (PHP 8.2) ─────────────────────────────────────
#[test]
fn readonly_class_deep() {
    assert_eq!(run_prints(r#"<?php
readonly class Coordinate {
    public function __construct(
        public float $lat,
        public float $lng,
    ) {}
    public function distanceTo(Coordinate $other): float {
        return sqrt(($this->lat - $other->lat) ** 2 + ($this->lng - $other->lng) ** 2);
    }
}
$a = new Coordinate(0, 0);
$b = new Coordinate(3, 4);
echo $b->lat;
echo $a->distanceTo($b);
"#), &["3", "5"]);
}

// ── Union types (deep) ───────────────────────────────────────────
#[test]
fn union_types_function() {
    assert_eq!(run_prints(r#"<?php
function stringify(int|float|string $val): string {
    if (is_int($val)) return "int:$val";
    if (is_float($val)) return "float:$val";
    return "str:$val";
}
echo stringify(42);
echo stringify(3.14);
echo stringify("hello");
"#), &["int:42", "float:3.14", "str:hello"]);
}

#[test]
fn nullable_type_function() {
    assert_eq!(run_prints(r#"<?php
function greet(?string $name): string {
    return $name !== null ? "Hello, $name" : "Hello, stranger";
}
echo greet("Alice");
echo greet(null);
"#), &["Hello, Alice", "Hello, stranger"]);
}

// ── Named arguments (deep) ───────────────────────────────────────
#[test]
fn named_args_with_spread() {
    assert_eq!(run_prints(r#"<?php
function formatDate(int $year, int $month, int $day): string {
    return sprintf("%04d-%02d-%02d", $year, $month, $day);
}
$params = ["month" => 6, "day" => 15, "year" => 2024];
echo formatDate(...$params);
"#), &["2024-06-15"]);
}

#[test]
fn named_args_in_builtin() {
    assert_eq!(run_prints(r#"<?php
echo implode(separator: ", ", array: ["a", "b", "c"]);
echo str_pad(string: "42", length: 5, pad_string: "0", pad_type: STR_PAD_LEFT);
"#), &["a, b, c", "00042"]);
}

// ── Nullsafe operator (deep) ─────────────────────────────────────
#[test]
fn nullsafe_deep_chain() {
    assert_eq!(run_prints(r#"<?php
class Country {
    public function __construct(public string $name) {}
}
class Address {
    public ?Country $country;
    public function __construct(?Country $country) {
        $this->country = $country;
    }
}
class User {
    public ?Address $address;
    public function __construct(?Address $address) {
        $this->address = $address;
    }
}
$u1 = new User(new Address(new Country("USA")));
$u2 = new User(new Address(null));
$u3 = new User(null);
echo $u1?->address?->country?->name ?? "unknown";
echo $u2?->address?->country?->name ?? "unknown";
echo $u3?->address?->country?->name ?? "unknown";
"#), &["USA", "unknown", "unknown"]);
}

#[test]
fn nullsafe_method_call() {
    assert_eq!(run_prints(r#"<?php
class Repo {
    public function find(int $id): ?string {
        return $id === 1 ? "found" : null;
    }
}
$repo = new Repo();
echo $repo->find(1)?->length ?? "null";
echo strlen($repo->find(1) ?? "");
echo $repo->find(99) ?? "not found";
"#), &["null", "5", "not found"]);
}

// ── Attributes (PHP 8.0) ─────────────────────────────────────────
#[test]
fn attributes_on_class_and_method() {
    assert_eq!(run_prints(r#"<?php
#[Attribute]
class Route {
    public function __construct(public string $path) {}
}
#[Route("/api/users")]
class UserController {
    #[Route("/list")]
    public function list(): string {
        return "user list";
    }
}
$c = new UserController();
echo $c->list();
"#), &["user list"]);
}

// ── Arrow functions (deep) ───────────────────────────────────────
#[test]
fn arrow_fn_complex_expressions() {
    assert_eq!(run_prints(r#"<?php
$add = fn(int $a, int $b): int => $a + $b;
$compose = fn(callable $f, callable $g) => fn($x) => $f($g($x));
$double = fn($x) => $x * 2;
$inc = fn($x) => $x + 1;
$doubleInc = $compose($double, $inc);
echo $add(3, 4);
echo $doubleInc(5);
"#), &["7", "12"]);
}

// ── Spread in function calls and arrays ──────────────────────────
#[test]
fn spread_in_various_contexts() {
    assert_eq!(run_prints(r#"<?php
function sum(int ...$nums): int {
    return array_sum($nums);
}
echo sum(1, 2, 3);
echo sum(...[4, 5, 6]);

$a = [1, 2, 3];
$b = [4, 5, 6];
$merged = [...$a, ...$b];
echo implode(",", $merged);
"#), &["6", "15", "1,2,3,4,5,6"]);
}

// ── Mixed type ───────────────────────────────────────────────────
#[test]
fn mixed_type_hint() {
    assert_eq!(run_prints(r#"<?php
function display(mixed $value): string {
    return match(gettype($value)) {
        "integer" => "int:$value",
        "string" => "str:$value",
        "array" => "arr:" . count($value),
        "NULL" => "null",
        default => "other",
    };
}
echo display(42);
echo display("hello");
echo display([1, 2, 3]);
echo display(null);
"#), &["int:42", "str:hello", "arr:3", "null"]);
}

// ── Never type ───────────────────────────────────────────────────
#[test]
fn never_return_type() {
    assert_eq!(run_prints(r#"<?php
function fail(string $msg): never {
    throw new RuntimeException($msg);
}
try {
    fail("fatal error");
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
"#), &["fatal error"]);
}

// ── Constant expressions in enums ────────────────────────────────
#[test]
fn enum_constant_expressions() {
    assert_eq!(run_prints(r#"<?php
enum Permission: int {
    case Read = 1;
    case Write = 2;
    case Execute = 4;
}
$perms = Permission::Read->value | Permission::Write->value;
echo $perms;
echo ($perms & Permission::Read->value) ? "can read" : "no read";
echo ($perms & Permission::Execute->value) ? "can exec" : "no exec";
"#), &["3", "can read", "no exec"]);
}

// ── Disjunctive Normal Form (DNF) types ──────────────────────────
#[test]
fn dnf_type_hint() {
    assert_eq!(run_prints(r#"<?php
interface Loggable {
    public function toLog(): string;
}
class Event implements Loggable {
    public function __construct(public string $name) {}
    public function toLog(): string { return "event:$this->name"; }
}
function process((Loggable&Stringable)|string $input): string {
    if (is_string($input)) return "str:$input";
    return $input->toLog();
}
echo process("hello");
"#), &["str:hello"]);
}

// ── Closure::bind and Closure::fromCallable ──────────────────────
#[test]
fn closure_bind_to_object() {
    assert_eq!(run_prints(r#"<?php
class Counter {
    private int $count = 0;
}
$increment = Closure::bind(function() {
    $this->count++;
    return $this->count;
}, new Counter(), Counter::class);
echo $increment();
echo $increment();
echo $increment();
"#), &["1", "2", "3"]);
}

#[test]
fn closure_from_callable() {
    assert_eq!(run_prints(r#"<?php
function double(int $x): int { return $x * 2; }
$fn = Closure::fromCallable('double');
echo $fn(5);
echo $fn(21);
"#), &["10", "42"]);
}
