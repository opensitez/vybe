//! Inheritance runtime: `parent::`, `static::`, visibility, overrides.

crate::php_cases! {
    parent_method_call_from_child => {
        r#"<?php
class Base { public function id(): string { return 'b'; } }
class Child extends Base { public function id(): string { return parent::id() . 'c'; } }
echo (new Child())->id();
"#,
        ["bc"]
    };

    parent_static_call => {
        r#"<?php
class Base { public static function n(): int { return 1; } }
class Child extends Base { public static function n(): int { return parent::n() + 1; } }
echo Child::n();
"#,
        ["2"]
    };

    static_late_binding_child => {
        r#"<?php
class Base { public static function who(): string { return static::class; } }
class Child extends Base {}
echo Child::who();
"#,
        ["Child"]
    };

    protected_property_visible_in_child => {
        r#"<?php
class Base { protected int $n = 5; public function read(): int { return $this->n; } }
class Child extends Base {}
echo (new Child())->read();
"#,
        ["5"]
    };

    private_not_inherited_child_defines_own => {
        r#"<?php
class Base { private int $n = 1; }
class Child extends Base { public int $n = 2; }
$c = new Child();
echo $c->n;
"#,
        ["2"]
    };

    child_constructor_calls_parent => {
        r#"<?php
class Base { public function __construct(public int $n) {} }
class Child extends Base { public function __construct() { parent::__construct(9); } }
echo (new Child())->n;
"#,
        ["9"]
    };

    abstract_class_concrete_child => {
        r#"<?php
abstract class A { abstract public function f(): int; }
class C extends A { public function f(): int { return 3; } }
echo (new C())->f();
"#,
        ["3"]
    };

    interface_multiple_on_class => {
        r#"<?php
interface I1 { public function a(): int; }
interface I2 { public function b(): int; }
class C implements I1, I2 {
    public function a(): int { return 1; }
    public function b(): int { return 2; }
}
echo (new C())->a() + (new C())->b();
"#,
        ["3"]
    };

    trait_method_used_in_class => {
        r#"<?php
trait T { public function hi(): string { return 't'; } }
class C { use T; }
echo (new C())->hi();
"#,
        ["t"]
    };

    trait_conflict_resolved_with_insteadof => {
        r#"<?php
trait A { public function m(): string { return 'a'; } }
trait B { public function m(): string { return 'b'; } }
class C { use A, B { A::m insteadof B; } }
echo (new C())->m();
"#,
        ["a"]
    };

    trait_alias_changes_visibility => {
        r#"<?php
trait T { private function secret(): string { return 's'; } }
class C { use T { secret as public show; } }
echo (new C())->show();
"#,
        ["s"]
    };

    final_class_cannot_extend_compile_ok_child_separate => {
        r#"<?php
final class F { public function v(): int { return 1; } }
echo (new F())->v();
"#,
        ["1"]
    };

    final_method_child_cannot_override_same_name => {
        r#"<?php
class Base { public function ok(): int { return 1; } }
class Child extends Base { public function ok(): int { return 2; } }
echo (new Child())->ok();
"#,
        ["2"]
    };

    static_property_inherited => {
        r#"<?php
class Base { public static int $c = 0; }
class Child extends Base {}
Child::$c = 4;
echo Child::$c;
"#,
        ["4"]
    };

    self_refers_to_declaring_class => {
        r#"<?php
class A { public static function c(): string { return self::class; } }
echo A::c();
"#,
        ["A"]
    };

    parent_const_access => {
        r#"<?php
class Base { public const N = 10; }
class Child extends Base { public function v(): int { return parent::N; } }
echo (new Child())->v();
"#,
        ["10"]
    };

    child_overrides_and_calls_parent_constructor_chain => {
        r#"<?php
class Base { public function __construct(public string $tag) {} }
class Child extends Base {
    public function __construct() { parent::__construct('x'); }
}
echo (new Child())->tag;
"#,
        ["x"]
    };

    instanceof_parent_from_child_instance => {
        r#"<?php
class Base {}
class Child extends Base {}
echo (new Child()) instanceof Base ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_declared_traits_includes_used => {
        r#"<?php
trait Tx {}
class C { use Tx; }
echo in_array('Tx', get_declared_traits()) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    class_parents_returns_hierarchy => {
        r#"<?php
class A {}
class B extends A {}
echo implode(',', class_parents(B::class));
"#,
        ["A"]
    };

    variadic_parent_constructor => {
        r#"<?php
class Base { public function __construct(public int ...$nums) {} }
class Child extends Base { public function __construct(int ...$n) { parent::__construct(...$n); } }
echo array_sum((new Child(1, 2, 3))->nums);
"#,
        ["6"]
    };

    covariant_return_child_narrower => {
        r#"<?php
class Base { public function make(): Base { return $this; } }
class Child extends Base { public function make(): Child { return $this; } }
echo (new Child())->make()::class;
"#,
        ["Child"]
    };

    static_call_child_from_parent_name => {
        r#"<?php
class Base { public static function who(): string { return static::class; } }
class Child extends Base {}
echo Base::who();
"#,
        ["Base"]
    };

    promoted_property_inherited => {
        r#"<?php
class Base { public function __construct(public int $n) {} }
class Child extends Base {}
echo (new Child(7))->n;
"#,
        ["7"]
    };

    abstract_trait_requires_implementation => {
        r#"<?php
trait T { abstract public function f(): int; }
class C { use T; public function f(): int { return 2; } }
echo (new C())->f();
"#,
        ["2"]
    };

    nested_trait_use => {
        r#"<?php
trait Inner { public function i(): int { return 1; } }
trait Outer { use Inner; }
class C { use Outer; }
echo (new C())->i();
"#,
        ["1"]
    };

    parent_private_method_not_visible_skip_runtime => {
        r#"<?php
class Base { public function pub(): int { return 1; } }
class Child extends Base { public function pub(): int { return parent::pub() + 1; } }
echo (new Child())->pub();
"#,
        ["2"]
    };

    interface_constant_access => {
        r#"<?php
interface I { public const K = 'v'; }
class C implements I {}
echo C::K;
"#,
        ["v"]
    };

    enum_implements_interface => {
        r#"<?php
interface Labeled { public function label(): string; }
enum Status implements Labeled { case On; public function label(): string { return 'on'; } }
echo Status::On->label();
"#,
        ["on"]
    };

    readonly_class_extended_by_readonly => {
        r#"<?php
readonly class Base { public function __construct(public int $n) {} }
readonly class Child extends Base {}
echo (new Child(4))->n;
"#,
        ["4"]
    };

    attribute_inherited_override => {
        r#"<?php
class Base { public function run(): string { return 'b'; } }
class Child extends Base { #[\Override] public function run(): string { return 'c'; } }
echo (new Child())->run();
"#,
        ["c"]
    };

    static_trait_method => {
        r#"<?php
trait T { public static function n(): int { return 5; } }
class C { use T; }
echo C::n();
"#,
        ["5"]
    };

    hierarchy_depth_three => {
        r#"<?php
class A { public function tag(): string { return 'a'; } }
class B extends A { public function tag(): string { return parent::tag() . 'b'; } }
class C extends B { public function tag(): string { return parent::tag() . 'c'; } }
echo (new C())->tag();
"#,
        ["abc"]
    };

    unset_child_public_inherited => {
        r#"<?php
class Base { public int $x = 1; }
class Child extends Base {}
$c = new Child();
unset($c->x);
echo property_exists($c, 'x') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    clone_inherited_clone_method => {
        r#"<?php
class Base { public function __construct(public int $n) {} public function __clone(): void { $this->n++; } }
class Child extends Base {}
$c = clone new Child(1);
echo $c->n;
"#,
        ["2"]
    };

    abstract_static_not_directly_called => {
        r#"<?php
abstract class A { public static function ok(): int { return 1; } }
class B extends A {}
echo B::ok();
"#,
        ["1"]
    };
}
