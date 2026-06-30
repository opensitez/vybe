//! Static properties, methods, constants, and late static binding.

crate::php_cases! {
    static_property_shared_across_instances => {
        r#"<?php
class Counter { public static int $n = 0; }
Counter::$n = 3;
echo Counter::$n;
"#,
        ["3"]
    };

    static_method_call_without_instance => {
        r#"<?php
class MathUtil { public static function double(int $x): int { return $x * 2; } }
echo MathUtil::double(4);
"#,
        ["8"]
    };

    static_local_in_function_persists => {
        r#"<?php
function seq(): int { static $i = 0; return ++$i; }
echo seq() . seq();
"#,
        ["12"]
    };

    self_keyword_in_static_method => {
        r#"<?php
class A { public static function who(): string { return self::class; } }
echo A::who();
"#,
        ["A"]
    };

    static_constant_access => {
        r#"<?php
class Config { public const MAX = 100; }
echo Config::MAX;
"#,
        ["100"]
    };

    parent_static_constant_in_child => {
        r#"<?php
class Base { public const K = 'b'; }
class Child extends Base {}
echo Child::K;
"#,
        ["b"]
    };

    late_static_binding_child_class => {
        r#"<?php
class Base { public static function id(): string { return static::class; } }
class Child extends Base {}
echo Child::id();
"#,
        ["Child"]
    };

    static_property_increment => {
        r#"<?php
class Hits { public static int $c = 0; public static function bump(): int { return ++self::$c; } }
echo Hits::bump() . Hits::bump();
"#,
        ["12"]
    };

    instance_reads_static_via_class => {
        r#"<?php
class Store { public static string $name = 'main'; }
echo Store::$name;
"#,
        ["main"]
    };

    static_trait_method => {
        r#"<?php
trait T { public static function tag(): string { return 't'; } }
class C { use T; }
echo C::tag();
"#,
        ["t"]
    };

    static_private_accessible_in_class => {
        r#"<?php
class Vault { private static int $secret = 9; public static function peek(): int { return self::$secret; } }
echo Vault::peek();
"#,
        ["9"]
    };

    static_array_property_mutation => {
        r#"<?php
class Registry { public static array $items = []; }
Registry::$items[] = 'a';
echo count(Registry::$items);
"#,
        ["1"]
    };

    static_return_new_instance => {
        r#"<?php
class Node { public static function make(): self { return new self(); } }
echo (new Node())::class;
"#,
        ["Node"]
    };

    static_closure_captures_static => {
        r#"<?php
class Logger {
    public static int $level = 1;
    public static function run(): int {
        return (function (): int { return self::$level; })();
    }
}
echo Logger::run();
"#,
        ["1"]
    };

    interface_constant_via_implementer => {
        r#"<?php
interface Limits { public const MAX = 5; }
class App implements Limits {}
echo App::MAX;
"#,
        ["5"]
    };

    enum_static_cases_count => {
        r#"<?php
enum Size { case S; case M; case L; }
echo count(Size::cases());
"#,
        ["3"]
    };

    static_nullable_property_default => {
        r#"<?php
class Cache { public static ?string $key = null; }
echo Cache::$key === null ? 'null' : 'set';
"#,
        ["null"]
    };

    static_method_calls_other_static => {
        r#"<?php
class A { public static function one(): int { return 1; } public static function two(): int { return self::one() + 1; } }
echo A::two();
"#,
        ["2"]
    };

    static_property_in_child_shadows_parent => {
        r#"<?php
class Base { public static int $n = 1; }
class Child extends Base { public static int $n = 2; }
echo Child::$n;
"#,
        ["2"]
    };

    static_method_inheritance_override => {
        r#"<?php
class Base { public static function v(): int { return 1; } }
class Child extends Base { public static function v(): int { return parent::v() + 1; } }
echo Child::v();
"#,
        ["2"]
    };

    class_constant_magic_const_class => {
        r#"<?php
class Demo { public function name(): string { return __CLASS__; } }
echo (new Demo())->name();
"#,
        ["Demo"]
    };

    static_var_in_method_not_function => {
        r#"<?php
class Tick { public function n(): int { static $c = 0; return ++$c; } }
$t = new Tick();
echo $t->n() . $t->n();
"#,
        ["12"]
    };

    constant_function_returns_static_value => {
        r#"<?php
define('APP_VER', '1.0');
echo constant('APP_VER');
"#,
        ["1.0"]
    };

    static_promoted_readonly_not_applicable_use_normal => {
        r#"<?php
class Point { public static int $x = 0; public static function set(int $v): void { self::$x = $v; } }
Point::set(4);
echo Point::$x;
"#,
        ["4"]
    };

    static_abstract_concrete => {
        r#"<?php
abstract class Base { public static function ok(): int { return 1; } }
class Child extends Base {}
echo Child::ok();
"#,
        ["1"]
    };
}
