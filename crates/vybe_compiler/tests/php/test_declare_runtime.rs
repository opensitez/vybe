//! `declare()` directives — strict types, encoding, ticks (runtime output).

crate::php_cases! {
    declare_strict_types_int_addition => {
        r#"<?php
declare(strict_types=1);
function add(int $a, int $b): int { return $a + $b; }
echo add(2, 3);
"#,
        ["5"]
    };

    declare_strict_types_string_upper => {
        r#"<?php
declare(strict_types=1);
function shout(string $s): string { return strtoupper($s); }
echo shout('hi');
"#,
        ["HI"]
    };

    declare_strict_types_bool_negation => {
        r#"<?php
declare(strict_types=1);
function flip(bool $b): bool { return !$b; }
echo flip(true) ? '1' : '0';
"#,
        ["0"]
    };

    declare_strict_types_float_multiply => {
        r#"<?php
declare(strict_types=1);
function dbl(float $x): float { return $x * 2.0; }
echo (int)dbl(2.5);
"#,
        ["5"]
    };

    declare_strict_types_array_count => {
        r#"<?php
declare(strict_types=1);
function len(array $a): int { return count($a); }
echo len([1, 2, 3]);
"#,
        ["3"]
    };

    declare_strict_types_nullable_param => {
        r#"<?php
declare(strict_types=1);
function show(?string $s): string { return $s ?? 'none'; }
echo show(null);
"#,
        ["none"]
    };

    declare_strict_types_union_int_string => {
        r#"<?php
declare(strict_types=1);
function tag(int|string $v): string { return (string)$v; }
echo tag(7);
"#,
        ["7"]
    };

    declare_strict_types_return_type_bool => {
        r#"<?php
declare(strict_types=1);
function is_even(int $n): bool { return $n % 2 === 0; }
echo is_even(4) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    declare_strict_types_promoted_constructor => {
        r#"<?php
declare(strict_types=1);
class Box { public function __construct(public int $n) {} }
echo (new Box(9))->n;
"#,
        ["9"]
    };

    declare_strict_types_void_function => {
        r#"<?php
declare(strict_types=1);
function log_msg(string $m): void { echo $m; }
log_msg('ok');
"#,
        ["ok"]
    };

    declare_encoding_utf8_string_length => {
        r#"<?php
declare(encoding='UTF-8');
echo strlen('é');
"#,
        ["2"]
    };

    declare_ticks_function_runs => {
        r#"<?php
declare(ticks=1);
$hits = 0;
register_tick_function(function () use (&$hits) { $hits++; });
for ($i = 0; $i < 3; $i++) {}
echo $hits >= 0 ? 'ticks' : 'no';
"#,
        ["ticks"]
    };

    declare_strict_in_namespace => {
        r#"<?php
namespace DeclTest;
declare(strict_types=1);
function id(int $n): int { return $n; }
echo id(4);
"#,
        ["4"]
    };

    declare_strict_static_return_type => {
        r#"<?php
declare(strict_types=1);
class Calc { public static function sq(int $n): int { return $n * $n; } }
echo Calc::sq(6);
"#,
        ["36"]
    };

    declare_strict_readonly_promoted => {
        r#"<?php
declare(strict_types=1);
readonly class Id { public function __construct(public int $v) {} }
echo (new Id(3))->v;
"#,
        ["3"]
    };

    declare_strict_enum_param => {
        r#"<?php
declare(strict_types=1);
enum Color { case Red; case Blue; }
function paint(Color $c): string { return $c->name; }
echo paint(Color::Red);
"#,
        ["Red"]
    };

    declare_strict_match_return => {
        r#"<?php
declare(strict_types=1);
function sign(int $n): string {
    return match (true) { $n < 0 => 'neg', $n > 0 => 'pos', default => 'zero' };
}
echo sign(0);
"#,
        ["zero"]
    };

    declare_strict_arrow_closure => {
        r#"<?php
declare(strict_types=1);
$inc = fn(int $n): int => $n + 1;
echo $inc(5);
"#,
        ["6"]
    };

    declare_strict_interface_impl => {
        r#"<?php
declare(strict_types=1);
interface Adder { public function add(int $a, int $b): int; }
class Plus implements Adder { public function add(int $a, int $b): int { return $a + $b; } }
echo (new Plus())->add(2, 2);
"#,
        ["4"]
    };

    declare_strict_trait_method => {
        r#"<?php
declare(strict_types=1);
trait T { public function twice(int $n): int { return $n * 2; } }
class C { use T; }
echo (new C())->twice(3);
"#,
        ["6"]
    };

    declare_strict_generators_yield => {
        r#"<?php
declare(strict_types=1);
function gen(): Generator { yield 1; yield 2; }
echo implode('', iterator_to_array(gen()));
"#,
        ["12"]
    };

    declare_strict_named_args => {
        r#"<?php
declare(strict_types=1);
function pair(int $a, int $b): int { return $a + $b; }
echo pair(b: 2, a: 3);
"#,
        ["5"]
    };

    declare_strict_attribute_class => {
        r#"<?php
declare(strict_types=1);
#[\AllowDynamicProperties]
class Flex {}
$f = new Flex();
$f->x = 1;
echo $f->x;
"#,
        ["1"]
    };

    declare_strict_finally_returns => {
        r#"<?php
declare(strict_types=1);
function run(): string {
    try { return 'try'; } finally { echo '!'; }
}
echo run();
"#,
        ["!try"]
    };

    declare_strict_multiple_in_file_last_wins_encoding => {
        r#"<?php
declare(encoding='UTF-8');
echo 'utf8';
"#,
        ["utf8"]
    };
}
