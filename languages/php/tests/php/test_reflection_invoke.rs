//! `ReflectionFunction::invoke`, `ReflectionMethod::invoke`, and constant fetch.

crate::php_cases! {
    reflectionfunction_invoke_calls_user_function => {
        r#"<?php
function add(int $a, int $b): int { return $a + $b; }
$ref = new ReflectionFunction('add');
echo $ref->invoke(2, 3);
"#,
        ["5"]
    };

    reflectionmethod_invoke_on_instance => {
        r#"<?php
class Calc { public function mul(int $x): int { return $x * 3; } }
$ref = new ReflectionMethod(Calc::class, 'mul');
echo $ref->invoke(new Calc(), 4);
"#,
        ["12"]
    };

    reflectionmethod_invoke_static => {
        r#"<?php
class Id { public static function val(): int { return 42; } }
$ref = new ReflectionMethod(Id::class, 'val');
echo $ref->invoke(null);
"#,
        ["42"]
    };

    reflectionclass_getconstant_value => {
        r#"<?php
class C { public const N = 99; }
$ref = new ReflectionClass(C::class);
echo $ref->getConstant('N');
"#,
        ["99"]
    };

    reflectionclass_hasmethod_detects_public => {
        r#"<?php
class S { public function go(): void {} }
$ref = new ReflectionClass(S::class);
echo $ref->hasMethod('go') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    reflectionparameter_isoptional_with_default => {
        r#"<?php
function f(int $a = 1): void {}
$ref = new ReflectionFunction('f');
echo $ref->getParameters()[0]->isOptional() ? 'yes' : 'no';
"#,
        ["yes"]
    };

    reflectionparameter_getdefaultvalue => {
        r#"<?php
function f(string $s = 'd'): void {}
$ref = new ReflectionFunction('f');
echo $ref->getParameters()[0]->getDefaultValue();
"#,
        ["d"]
    };

    reflectionclass_isinstantiable_for_concrete => {
        r#"<?php
class X {}
$ref = new ReflectionClass(X::class);
echo $ref->isInstantiable() ? 'yes' : 'no';
"#,
        ["yes"]
    };

    reflectionclass_isabstract_for_abstract => {
        r#"<?php
abstract class A {}
$ref = new ReflectionClass(A::class);
echo $ref->isAbstract() ? 'yes' : 'no';
"#,
        ["yes"]
    };

    reflectionenum_getcases_count => {
        r#"<?php
enum Color { case Red; case Blue; }
$ref = new ReflectionEnum(Color::class);
echo count($ref->getCases());
"#,
        ["2"]
    };

    constant_fetch_class_constant => {
        r#"<?php
class K { public const V = 'ok'; }
echo K::V;
"#,
        ["ok"]
    };

    defined_checks_constant_exists => {
        r#"<?php
echo defined('PHP_VERSION') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    reflectionfunction_returns_reference_false_for_add => {
        r#"<?php
function add(int $a, int $b): int { return $a + $b; }
$ref = new ReflectionFunction('add');
echo $ref->returnsReference() ? 'yes' : 'no';
"#,
        ["no"]
    };
}
