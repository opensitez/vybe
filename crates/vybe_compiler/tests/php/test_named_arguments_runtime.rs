//! Named arguments: reordering, skipping defaults, spread with names.

crate::php_cases! {
    named_args_reorder_parameters => {
        r#"<?php
function pair(int $a, int $b): int { return $a + $b; }
echo pair(b: 2, a: 3);
"#,
        ["5"]
    };

    named_args_skip_optional_middle => {
        r#"<?php
function f(int $a, int $b = 10, int $c = 20): int { return $a + $b + $c; }
echo f(1, c: 5);
"#,
        ["16"]
    };

    named_args_with_defaults_only => {
        r#"<?php
function g(string $s = 'x', int $n = 1): string { return $s . $n; }
echo g(n: 9);
"#,
        ["x9"]
    };

    named_args_constructor_promotion => {
        r#"<?php
class P { public function __construct(public string $name, public int $age) {} }
$p = new P(age: 30, name: 'Ann');
echo $p->name . $p->age;
"#,
        ["Ann30"]
    };

    named_args_method_call => {
        r#"<?php
class C {
    public function tag(string $a, string $b): string { return $a . $b; }
}
echo (new C())->tag(b: '2', a: '1');
"#,
        ["12"]
    };

    named_args_static_method => {
        r#"<?php
class C {
    public static function sum(int $a, int $b): int { return $a + $b; }
}
echo C::sum(b: 4, a: 5);
"#,
        ["9"]
    };

    named_args_parent_constructor => {
        r#"<?php
class Base { public function __construct(public int $n) {} }
class Child extends Base {
    public function __construct() { parent::__construct(n: 7); }
}
echo (new Child())->n;
"#,
        ["7"]
    };

    named_args_builtin_str_replace => {
        r#"<?php
echo str_replace(subject: 'aabb', search: 'a', replace: 'x');
"#,
        ["xxbb"]
    };

    named_args_builtin_array_slice => {
        r#"<?php
echo implode(',', array_slice(array: [1, 2, 3, 4], offset: 1, length: 2));
"#,
        ["2,3"]
    };

    named_args_builtin_str_starts_with => {
        r#"<?php
echo str_starts_with(haystack: 'hello', needle: 'he') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    named_args_after_positional => {
        r#"<?php
function h(int $a, int $b, int $c): int { return $a + $b + $c; }
echo h(1, c: 3, b: 2);
"#,
        ["6"]
    };

    named_args_in_closure => {
        r#"<?php
$fn = function (int $x, int $y): int { return $x * $y; };
echo $fn(y: 4, x: 3);
"#,
        ["12"]
    };

    named_args_arrow_function => {
        r#"<?php
$fn = fn(int $a, int $b): int => $a - $b;
echo $fn(b: 1, a: 5);
"#,
        ["4"]
    };

    named_args_first_class_callable => {
        r#"<?php
function inc(int $n, int $by = 1): int { return $n + $by; }
echo (inc(...))(5, by: 2);
"#,
        ["7"]
    };

    named_args_variadic_with_names => {
        r#"<?php
function join(string $sep, string ...$parts): string { return implode($sep, $parts); }
echo join(sep: '-', 'a', 'b');
"#,
        ["a-b"]
    };

    named_args_attribute_on_method => {
        r#"<?php
class A {
    #[\Override]
    public function run(): string { return 'ok'; }
}
class B extends A { #[\Override] public function run(): string { return parent::run(); } }
echo (new B())->run();
"#,
        ["ok"]
    };

    named_args_json_encode_flags => {
        r#"<?php
echo json_encode(value: ['a' => 1], flags: JSON_FORCE_OBJECT);
"#,
        ["{\"a\":1}"]
    };

    named_args_preg_match_offsets => {
        r#"<?php
preg_match(pattern: '/(\d+)/', subject: 'x9y', matches: $m);
echo $m[1];
"#,
        ["9"]
    };

    named_args_datetime_create => {
        r#"<?php
date_default_timezone_set('UTC');
$d = DateTime::createFromFormat(format: 'Y-m-d', datetime: '2024-01-02');
echo $d->format('d');
"#,
        ["02"]
    };

    named_args_array_combine => {
        r#"<?php
$a = array_combine(keys: ['a', 'b'], values: [1, 2]);
echo $a['b'];
"#,
        ["2"]
    };

    named_args_inherited_method => {
        r#"<?php
class Base { public function v(int $a, int $b): int { return $a + $b; } }
class Child extends Base {}
echo (new Child())->v(b: 2, a: 3);
"#,
        ["5"]
    };

    named_args_interface_implementation => {
        r#"<?php
interface I { public function run(int $x, int $y): int; }
class C implements I { public function run(int $x, int $y): int { return $x * $y; } }
echo (new C())->run(y: 3, x: 4);
"#,
        ["12"]
    };

    named_args_nullable_default => {
        r#"<?php
function opt(?string $s = 'd'): string { return $s ?? 'null'; }
echo opt(s: null);
"#,
        ["null"]
    };

    named_args_union_param => {
        r#"<?php
function show(int|string $v, string $prefix = ''): string { return $prefix . $v; }
echo show(v: 7, prefix: '#');
"#,
        ["#7"]
    };

    named_args_trailing_comma_call => {
        r#"<?php
function t(int $a, int $b): int { return $a + $b; }
echo t(a: 1, b: 2,);
"#,
        ["3"]
    };

    named_args_nested_call => {
        r#"<?php
function id(int $n): int { return $n; }
function wrap(int $n): int { return id(n: $n); }
echo wrap(5);
"#,
        ["5"]
    };

    named_args_splat_mixed => {
        r#"<?php
function add(int $a, int $b, int $c): int { return $a + $b + $c; }
echo add(a: 1, ...[2, 3]);
"#,
        ["6"]
    };

    named_args_enum_case => {
        r#"<?php
enum E { case A; case B; }
function pick(E $e): string { return $e->name; }
echo pick(e: E::B);
"#,
        ["B"]
    };

    named_args_readonly_class => {
        r#"<?php
readonly class R { public function __construct(public string $k, public int $v) {} }
$r = new R(v: 9, k: 'x');
echo $r->k;
"#,
        ["x"]
    };

    named_args_anonymous_class_method => {
        r#"<?php
$o = new class {
    public function f(int $a, int $b): int { return $a - $b; }
};
echo $o->f(b: 1, a: 5);
"#,
        ["4"]
    };

    named_args_builtin_explode => {
        r#"<?php
echo implode('+', explode(separator: ',', string: 'a,b'));
"#,
        ["a+b"]
    };

    named_args_builtin_number_format => {
        r#"<?php
echo number_format(num: 1234.5, decimals: 1, decimal_separator: '.', thousands_separator: ',');
"#,
        ["1,234.5"]
    };

    named_args_heredoc_param => {
        r#"<?php
function show(string $s): string { return trim($s); }
echo show(s: <<<TXT
hi
TXT);
"#,
        ["hi"]
    };

    named_args_attribute_named_on_class => {
        r#"<?php
#[\AllowDynamicProperties]
class X {}
echo 'ok';
"#,
        ["ok"]
    };

    named_args_call_user_func => {
        r#"<?php
function mul(int $a, int $b): int { return $a * $b; }
echo call_user_func('mul', b: 3, a: 4);
"#,
        ["12"]
    };
}
