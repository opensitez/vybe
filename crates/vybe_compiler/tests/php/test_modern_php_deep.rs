use super::helpers::{compile_ok, run_prints};

// ── Match expression ─────────────────────────────────────────────

#[test]
fn match_with_complex_arms() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["ABF"]
    );
}

#[test]
fn match_multiple_conditions() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["successredirectnot foundserver errorunknown"]
    );
}

#[test]
fn match_no_default_throws() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    $x = 5;
    $result = match($x) {
        1 => "one",
        2 => "two",
    };
} catch (\UnhandledMatchError $e) {
    echo "unhandled";
}
"#
        ),
        &["unhandled"]
    );
}

#[test]
fn match_strict_comparison_no_coercion() {
    assert_eq!(
        run_prints(
            r#"<?php
$val = "0";
$result = match($val) {
    0   => "int zero",
    "0" => "string zero",
    default => "other",
};
echo $result;
"#
        ),
        &["string zero"]
    );
}

#[test]
fn match_returning_complex_value() {
    assert_eq!(
        run_prints(
            r#"<?php
function getConfig(string $env): array {
    return match($env) {
        "dev"  => ["debug" => true,  "log" => "verbose"],
        "prod" => ["debug" => false, "log" => "error"],
        default => ["debug" => false, "log" => "warning"],
    };
}
$cfg = getConfig("dev");
echo $cfg["log"];
$cfg2 = getConfig("prod");
echo $cfg2["debug"] ? "debug" : "no-debug";
"#
        ),
        &["verboseno-debug"]
    );
}

#[test]
fn match_as_function_argument() {
    assert_eq!(
        run_prints(
            r#"<?php
function repeat(string $s, int $n): string { return str_repeat($s, $n); }
$x = 3;
echo repeat(match($x) { 1 => "a", 2 => "b", 3 => "c", default => "?" }, 4);
"#
        ),
        &["cccc"]
    );
}

// ── Named arguments ──────────────────────────────────────────────

#[test]
fn named_args_in_user_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function greet(string $name, string $greeting = "Hello"): string {
    return "$greeting, $name!";
}
echo greet(name: "Alice");
echo greet(name: "Bob", greeting: "Hi");
"#
        ),
        &["Hello, Alice!Hi, Bob!"]
    );
}

#[test]
fn named_args_skipping_defaults() {
    assert_eq!(
        run_prints(
            r#"<?php
function create(string $type, string $color = "black", int $size = 10): string {
    return "$color $type (size $size)";
}
echo create(type: "circle", size: 20);
"#
        ),
        &["black circle (size 20)"]
    );
}

#[test]
fn named_args_in_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    public function __construct(
        public string $host = "localhost",
        public int    $port = 3306,
        public string $db   = "app"
    ) {}
}
$c = new Config(port: 5432, db: "mydb");
echo $c->host;
echo $c->port;
echo $c->db;
"#
        ),
        &["localhost5432mydb"]
    );
}

#[test]
fn named_args_mixed_with_positional() {
    assert_eq!(
        run_prints(
            r#"<?php
function slice(array $arr, int $offset, ?int $length = null, bool $preserve = false): array {
    return array_slice($arr, $offset, $length, $preserve);
}
$a = [1, 2, 3, 4, 5];
$r = slice($a, 1, length: 3);
echo implode(",", $r);
"#
        ),
        &["2,3,4"]
    );
}

#[test]
fn named_args_in_builtin() {
    assert_eq!(
        run_prints(
            r#"<?php
echo implode(separator: ", ", array: ["a", "b", "c"]);
echo str_pad(string: "42", length: 5, pad_string: "0", pad_type: STR_PAD_LEFT);
"#
        ),
        &["a, b, c00042"]
    );
}

#[test]
fn named_args_with_spread() {
    assert_eq!(
        run_prints(
            r#"<?php
function formatDate(int $year, int $month, int $day): string {
    return sprintf("%04d-%02d-%02d", $year, $month, $day);
}
$params = ["month" => 6, "day" => 15, "year" => 2024];
echo formatDate(...$params);
"#
        ),
        &["2024-06-15"]
    );
}

#[test]
fn named_args_in_arrow_function() {
    compile_ok(
        r#"<?php
$pad = fn(string $s, int $len) => str_pad(string: $s, length: $len, pad_string: "-", pad_type: STR_PAD_BOTH);
echo $pad("hi", 8);
"#,
    );
}

// ── Nullsafe operator ─────────────────────────────────────────────

#[test]
fn nullsafe_chain_returning_null() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public ?string $email;
    public function __construct(?string $email) { $this->email = $email; }
    public function getEmail(): ?string { return $this->email; }
}
$u = new User(null);
$result = $u?->getEmail() ?? "no email";
echo $result;
"#
        ),
        &["no email"]
    );
}

#[test]
fn nullsafe_chain_succeeding() {
    assert_eq!(
        run_prints(
            r#"<?php
class Address {
    public function __construct(public string $city) {}
    public function getCity(): string { return $this->city; }
}
class User {
    public ?Address $address;
    public function __construct(?Address $addr) { $this->address = $addr; }
}
$u = new User(new Address("Paris"));
echo $u?->address?->getCity() ?? "unknown";
"#
        ),
        &["Paris"]
    );
}

#[test]
fn nullsafe_deep_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["USAunknownunknown"]
    );
}

#[test]
fn nullsafe_with_method_call_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class Repo {
    public function find(int $id): ?string {
        return $id === 1 ? "found" : null;
    }
}
$repo = new Repo();
echo strlen($repo->find(1) ?? "");
echo $repo->find(99) ?? "not found";
"#
        ),
        &["5not found"]
    );
}

#[test]
fn nullsafe_mixed_with_regular() {
    assert_eq!(
        run_prints(
            r#"<?php
class Tree {
    public ?Tree $left  = null;
    public ?Tree $right = null;
    public function __construct(public int $value) {}
}
$root = new Tree(1);
$root->left = new Tree(2);
echo $root->left?->value;
echo $root->right?->value ?? "null";
"#
        ),
        &["2null"]
    );
}

#[test]
fn nullsafe_combined_with_null_coalescing() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function get(string $key): ?Config {
        return isset($this->data[$key]) ? new Config($this->data[$key]) : null;
    }
    public function value(string $key): mixed { return $this->data[$key] ?? null; }
}
$cfg = new Config(["db" => ["host" => "localhost"]]);
echo $cfg->get("db")?->value("host") ?? "default";
echo $cfg->get("missing")?->value("host") ?? "default";
"#
        ),
        &["localhostdefault"]
    );
}

// ── Union types ───────────────────────────────────────────────────

#[test]
fn union_types_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function stringify(int|float|string $val): string {
    if (is_int($val)) return "int:$val";
    if (is_float($val)) return "float:$val";
    return "str:$val";
}
echo stringify(42);
echo stringify(3.14);
echo stringify("hello");
"#
        ),
        &["int:42float:3.14str:hello"]
    );
}

#[test]
fn union_type_in_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Variant {
    public int|string $value;
    public function __construct(int|string $v) { $this->value = $v; }
    public function type(): string { return is_int($this->value) ? "int" : "string"; }
}
$a = new Variant(42);
$b = new Variant("hello");
echo $a->type();
echo $b->type();
"#
        ),
        &["intstring"]
    );
}

#[test]
fn nullable_type_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function greet(?string $name): string {
    return $name !== null ? "Hello, $name" : "Hello, stranger";
}
echo greet("Alice");
echo greet(null);
"#
        ),
        &["Hello, AliceHello, stranger"]
    );
}

// ── Readonly properties ───────────────────────────────────────────

#[test]
fn readonly_property_in_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
class ImmutablePoint {
    public readonly float $x;
    public readonly float $y;
    public function __construct(float $x, float $y) {
        $this->x = $x;
        $this->y = $y;
    }
    public function distanceTo(ImmutablePoint $other): float {
        return sqrt(($this->x - $other->x) ** 2 + ($this->y - $other->y) ** 2);
    }
}
$a = new ImmutablePoint(0, 0);
$b = new ImmutablePoint(3, 4);
echo $b->x;
echo $a->distanceTo($b);
"#
        ),
        &["35"]
    );
}

#[test]
fn readonly_via_constructor_promotion() {
    assert_eq!(
        run_prints(
            r#"<?php
class Money {
    public function __construct(
        public readonly int    $amount,
        public readonly string $currency
    ) {}
    public function format(): string {
        return $this->amount . " " . $this->currency;
    }
}
$m = new Money(100, "USD");
echo $m->amount;
echo $m->currency;
echo $m->format();
"#
        ),
        &["100USD100 USD"]
    );
}

#[test]
fn readonly_class_deep() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["35"]
    );
}

// ── Fibers ────────────────────────────────────────────────────────

#[test]
fn fiber_basic_start_suspend() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function() {
    echo "start";
    Fiber::suspend();
    echo "end";
});
$fiber->start();
echo "between";
$fiber->resume();
"#
        ),
        &["startbetweenend"]
    );
}

#[test]
fn fiber_passing_value_to_suspend() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): string {
    $a = Fiber::suspend("first");
    $b = Fiber::suspend("second");
    return "result: $a + $b";
});
echo $fiber->start();
echo $fiber->resume("hello");
echo $fiber->resume("world");
echo $fiber->getReturn();
"#
        ),
        &["firstsecondresult: hello + world"]
    );
}

#[test]
fn fiber_getting_value_from_resume() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    $val = Fiber::suspend("waiting");
    echo "got: $val";
});
$suspended = $fiber->start();
echo $suspended;
$fiber->resume("ping");
"#
        ),
        &["waitinggot: ping"]
    );
}

#[test]
fn fiber_terminated_state_check() {
    assert_eq!(
        run_prints(
            r#"<?php
$fiber = new Fiber(function(): void {
    Fiber::suspend();
});
$fiber->start();
echo $fiber->isSuspended() ? "suspended" : "not suspended";
echo $fiber->isTerminated() ? "terminated" : "not terminated";
$fiber->resume();
echo $fiber->isTerminated() ? "terminated" : "not terminated";
"#
        ),
        &["suspendednot terminatedterminated"]
    );
}

// ── Enums ─────────────────────────────────────────────────────────

#[test]
fn enum_cases_iteration() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["4NorthWest"]
    );
}

#[test]
fn backed_enum_from_try_from() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status: string {
    case Active   = "active";
    case Inactive = "inactive";
    case Banned   = "banned";
}
$s = Status::from("active");
echo $s->name;
$t = Status::tryFrom("unknown");
echo $t === null ? "null" : $t->name;
"#
        ),
        &["Activenull"]
    );
}

#[test]
fn enum_implements_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["$€gbp"]
    );
}

#[test]
fn enum_with_trait() {
    assert_eq!(
        run_prints(
            r#"<?php
trait HasLabel {
    public function label(): string {
        return ucfirst(strtolower($this->name));
    }
}
enum Color {
    use HasLabel;
    case Red;
    case Green;
    case Blue;
}
echo Color::Red->label();
echo Color::Green->label();
"#
        ),
        &["RedGreen"]
    );
}

#[test]
fn enum_used_in_match() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Suit: string {
    case Hearts   = "H";
    case Diamonds = "D";
    case Clubs    = "C";
    case Spades   = "S";
}
function color(Suit $s): string {
    return match($s) {
        Suit::Hearts, Suit::Diamonds => "red",
        Suit::Clubs, Suit::Spades   => "black",
    };
}
echo color(Suit::Hearts);
echo color(Suit::Spades);
"#
        ),
        &["redblack"]
    );
}

#[test]
fn enum_used_as_array_key_via_value() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Role: string {
    case Admin = "admin";
    case User  = "user";
    case Guest = "guest";
}
$permissions = [
    Role::Admin->value => ["read", "write", "delete"],
    Role::User->value  => ["read", "write"],
    Role::Guest->value => ["read"],
];
$role = Role::User;
echo count($permissions[$role->value]);
echo $permissions[Role::Guest->value][0];
"#
        ),
        &["2read"]
    );
}

#[test]
fn enum_constant() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Permission: int {
    case Read    = 1;
    case Write   = 2;
    case Execute = 4;
    const int ALL = 7;
}
echo Permission::Read->value;
echo Permission::ALL;
$perms = Permission::Read->value | Permission::Write->value;
echo $perms;
"#
        ),
        &["173"]
    );
}

#[test]
fn enum_in_type_hint() {
    compile_ok(
        r#"<?php
enum Season { case Spring; case Summer; case Autumn; case Winter; }
function describe(Season $s): string {
    return match($s) {
        Season::Spring => "flowers",
        Season::Summer => "sun",
        Season::Autumn => "leaves",
        Season::Winter => "snow",
    };
}
echo describe(Season::Winter);
"#,
    );
}

// ── First-class callable syntax ───────────────────────────────────

#[test]
fn first_class_callable_builtin() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = strlen(...);
echo $fn("hello");
echo $fn("hi");
"#
        ),
        &["52"]
    );
}

#[test]
fn first_class_callable_instance_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Formatter {
    public function upper(string $s): string {
        return strtoupper($s);
    }
}
$f = new Formatter();
$fn = $f->upper(...);
echo $fn("hello");
"#
        ),
        &["HELLO"]
    );
}

#[test]
fn first_class_callable_static_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Math {
    public static function double(int $x): int { return $x * 2; }
}
$fn = Math::double(...);
echo $fn(5);
echo $fn(21);
"#
        ),
        &["1042"]
    );
}

#[test]
fn first_class_callable_in_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$words = ["hello", "world", "php"];
$result = array_map(strtoupper(...), $words);
echo implode(",", $result);
"#
        ),
        &["HELLO,WORLD,PHP"]
    );
}

// ── Intersection types ────────────────────────────────────────────

#[test]
fn intersection_type_in_param() {
    compile_ok(
        r#"<?php
interface Countable2 { public function size(): int; }
interface Iterable2  { public function items(): array; }
class Bag implements Countable2, Iterable2 {
    private array $data;
    public function __construct(array $d) { $this->data = $d; }
    public function size(): int { return count($this->data); }
    public function items(): array { return $this->data; }
}
function process(Countable2&Iterable2 $obj): string {
    return "size=" . $obj->size() . ",items=" . count($obj->items());
}
echo process(new Bag([1, 2, 3]));
"#,
    );
}

// ── Never return type ─────────────────────────────────────────────

#[test]
fn never_return_type() {
    assert_eq!(
        run_prints(
            r#"<?php
function fail(string $msg): never {
    throw new RuntimeException($msg);
}
try {
    fail("fatal error");
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
"#
        ),
        &["fatal error"]
    );
}

// ── PHP 8.2 Features ─────────────────────────────────────────────

#[test]
fn dnf_type_hint() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["str:hello"]
    );
}

// ── PHP 8.3 Features ─────────────────────────────────────────────

#[test]
fn typed_class_constants() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    const int    MAX_SIZE    = 1024;
    const string DEFAULT_ENV = "production";
    const float  TAX_RATE    = 0.08;
    const bool   DEBUG       = false;
}
echo Config::MAX_SIZE;
echo Config::DEFAULT_ENV;
"#
        ),
        &["1024production"]
    );
}

#[test]
fn dynamic_class_constant_fetch() {
    assert_eq!(
        run_prints(
            r#"<?php
class HttpStatus {
    const OK        = 200;
    const NOT_FOUND = 404;
    const ERROR     = 500;
}
$const = "NOT_FOUND";
echo HttpStatus::{$const};
"#
        ),
        &["404"]
    );
}

// ── PHP 8.4 Features ─────────────────────────────────────────────

#[test]
fn property_hook_structural() {
    compile_ok(
        r#"<?php
class Temperature {
    public float $celsius {
        get { return $this->celsius; }
        set(float $value) { $this->celsius = $value; }
    }
    public float $fahrenheit {
        get { return $this->celsius * 9/5 + 32; }
    }
}
$t = new Temperature();
$t->celsius = 100.0;
echo $t->fahrenheit;
"#,
    );
}

#[test]
fn asymmetric_visibility_structural() {
    compile_ok(
        r#"<?php
class Counter {
    public private(set) int $count = 0;
    public function increment(): void { $this->count++; }
}
$c = new Counter();
$c->increment();
$c->increment();
echo $c->count;
"#,
    );
}

// ── Closure advanced ─────────────────────────────────────────────

#[test]
fn closure_bind_to_object() {
    assert_eq!(
        run_prints(
            r#"<?php
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
"#
        ),
        &["123"]
    );
}

#[test]
fn closure_from_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
function double(int $x): int { return $x * 2; }
$fn = Closure::fromCallable('double');
echo $fn(5);
echo $fn(21);
"#
        ),
        &["1042"]
    );
}

// ── Arrow functions deep ──────────────────────────────────────────

#[test]
fn arrow_fn_complex_expressions() {
    assert_eq!(
        run_prints(
            r#"<?php
$add     = fn(int $a, int $b): int => $a + $b;
$compose = fn(callable $f, callable $g) => fn($x) => $f($g($x));
$double  = fn($x) => $x * 2;
$inc     = fn($x) => $x + 1;
$doubleInc = $compose($double, $inc);
echo $add(3, 4);
echo $doubleInc(5);
"#
        ),
        &["712"]
    );
}

// ── Spread operator ───────────────────────────────────────────────

#[test]
fn spread_in_various_contexts() {
    assert_eq!(
        run_prints(
            r#"<?php
function sum(int ...$nums): int {
    return array_sum($nums);
}
echo sum(1, 2, 3);
echo sum(...[4, 5, 6]);

$a = [1, 2, 3];
$b = [4, 5, 6];
$merged = [...$a, ...$b];
echo implode(",", $merged);
"#
        ),
        &["6151,2,3,4,5,6"]
    );
}

// ── Mixed type hint ───────────────────────────────────────────────

#[test]
fn mixed_type_hint() {
    assert_eq!(
        run_prints(
            r#"<?php
function display(mixed $value): string {
    return match(gettype($value)) {
        "integer" => "int:$value",
        "string"  => "str:$value",
        "array"   => "arr:" . count($value),
        "NULL"    => "null",
        default   => "other",
    };
}
echo display(42);
echo display("hello");
echo display([1, 2, 3]);
echo display(null);
"#
        ),
        &["int:42str:helloarr:3null"]
    );
}

// ── Enum constant expressions ─────────────────────────────────────

#[test]
fn enum_constant_expressions() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Permission: int {
    case Read    = 1;
    case Write   = 2;
    case Execute = 4;
}
$perms = Permission::Read->value | Permission::Write->value;
echo $perms;
echo ($perms & Permission::Read->value)    ? "can read"  : "no read";
echo ($perms & Permission::Execute->value) ? "can exec"  : "no exec";
"#
        ),
        &["3can readno exec"]
    );
}

// ── Fiber scheduler pattern ───────────────────────────────────────

#[test]
fn fiber_scheduler_round_robin() {
    assert_eq!(
        run_prints(
            r#"<?php
$fibers = [];
for ($i = 1; $i <= 3; $i++) {
    $n = $i;
    $fibers[] = new Fiber(function() use ($n) {
        echo "task$n start";
        Fiber::suspend();
        echo "task$n end";
    });
}
foreach ($fibers as $f) { $f->start(); }
foreach ($fibers as $f) { $f->resume(); }
"#
        ),
        &["task1 starttask2 starttask3 starttask1 endtask2 endtask3 end"]
    );
}
