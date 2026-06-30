//! OOP runtime: `clone`, `isset`/`empty` on properties, comparisons, destruct order.

crate::php_cases! {
    clone_creates_shallow_copy_of_object => {
        r#"<?php
class Box { public function __construct(public int $n) {} }
$a = new Box(1);
$b = clone $a;
$b->n = 2;
echo $a->n . $b->n;
"#,
        ["12"]
    };

    clone_triggers_clone_method => {
        r#"<?php
class Tag {
    public function __construct(public string $v) {}
    public function __clone(): void { $this->v .= '!'; }
}
$t = new Tag('x');
$c = clone $t;
echo $c->v;
"#,
        ["x!"]
    };

    isset_object_property_true_when_set => {
        r#"<?php
class P { public ?int $n = 5; }
$p = new P();
echo isset($p->n) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    isset_object_property_false_when_null => {
        r#"<?php
class P { public ?int $n = null; }
$p = new P();
echo isset($p->n) ? 'yes' : 'no';
"#,
        ["no"]
    };

    empty_object_property_null_is_empty => {
        r#"<?php
class P { public ?string $s = null; }
$p = new P();
echo empty($p->s) ? 'empty' : 'set';
"#,
        ["empty"]
    };

    empty_object_property_zero_string_not_empty => {
        r#"<?php
class P { public string $s = '0'; }
$p = new P();
echo empty($p->s) ? 'empty' : 'set';
"#,
        ["set"]
    };

    object_identity_same_instance => {
        r#"<?php
class U {}
$a = new U();
echo $a === $a ? 'same' : 'diff';
"#,
        ["same"]
    };

    object_identity_different_instances => {
        r#"<?php
class U {}
echo (new U()) === (new U()) ? 'same' : 'diff';
"#,
        ["diff"]
    };

    spl_object_id_differs_per_instance => {
        r#"<?php
class U {}
echo spl_object_id(new U()) === spl_object_id(new U()) ? 'eq' : 'ne';
"#,
        ["ne"]
    };

    spl_object_hash_differs_per_instance => {
        r#"<?php
class U {}
echo spl_object_hash(new U()) === spl_object_hash(new U()) ? 'eq' : 'ne';
"#,
        ["ne"]
    };

    instanceof_checks_hierarchy => {
        r#"<?php
class A {}
class B extends A {}
echo (new B()) instanceof A ? 'yes' : 'no';
"#,
        ["yes"]
    };

    instanceof_false_for_unrelated => {
        r#"<?php
class A {}
class B {}
echo (new A()) instanceof B ? 'yes' : 'no';
"#,
        ["no"]
    };

    property_exists_public => {
        r#"<?php
class C { public int $x = 1; }
echo property_exists(new C(), 'x') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    property_exists_dynamic_after_set => {
        r#"<?php
class C {}
$o = new C();
$o->dyn = 1;
echo property_exists($o, 'dyn') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_object_vars_returns_public_properties => {
        r#"<?php
class C { public int $a = 1; private int $b = 2; }
echo implode(',', array_keys(get_object_vars(new C())));
"#,
        ["a"]
    };

    get_class_without_object_uses_this => {
        r#"<?php
class C {
    public static function name(): string { return get_class(); }
}
echo C::name();
"#,
        ["C"]
    };

    get_parent_class_returns_base => {
        r#"<?php
class Base {}
class Child extends Base {}
echo get_parent_class(new Child());
"#,
        ["Base"]
    };

    is_subclass_of_child => {
        r#"<?php
class Base {}
class Child extends Base {}
echo is_subclass_of('Child', 'Base') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_a_with_string_class_name => {
        r#"<?php
class Base {}
class Child extends Base {}
echo is_a(new Child(), Base::class) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    class_implements_interface_list => {
        r#"<?php
interface I { public function f(): void; }
class C implements I { public function f(): void {} }
echo implode('', class_implements(new C()));
"#,
        ["I"]
    };

    class_uses_trait_on_class => {
        r#"<?php
trait T { public function t(): int { return 1; } }
class C { use T; }
echo implode('', class_uses(C::class));
"#,
        ["T"]
    };

    destructor_runs_at_end_of_scope => {
        r#"<?php
class D { public function __destruct() { echo 'd'; } }
{ new D(); }
echo 'e';
"#,
        ["de"]
    };

    destructor_order_reverse_construction => {
        r#"<?php
class D {
    public function __construct(public string $k) {}
    public function __destruct() { echo $this->k; }
}
$a = new D('a');
$b = new D('b');
"#,
        ["ba"]
    };

    magic_get_returns_dynamic => {
        r#"<?php
class M {
    private array $d = ['k' => 'v'];
    public function __get(string $n) { return $this->d[$n] ?? null; }
}
echo (new M())->k;
"#,
        ["v"]
    };

    magic_set_stores_dynamic => {
        r#"<?php
class M {
    public array $d = [];
    public function __set(string $n, mixed $v): void { $this->d[$n] = $v; }
    public function read(): string { return $this->d['x']; }
}
$m = new M();
$m->x = 'ok';
echo $m->read();
"#,
        ["ok"]
    };

    magic_call_forwards => {
        r#"<?php
class M {
    public function __call(string $n, array $a): string { return $n . count($a); }
}
echo (new M())->foo(1, 2);
"#,
        ["foo2"]
    };

    magic_isset_on_container => {
        r#"<?php
class M {
    public function __construct(private array $d) {}
    public function __isset(string $k): bool { return isset($this->d[$k]); }
}
echo isset((new M(['a' => 1]))->a) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    magic_unset_clears => {
        r#"<?php
class M {
    public array $d = ['a' => 1];
    public function __unset(string $k): void { unset($this->d[$k]); }
}
$m = new M();
unset($m->a);
echo isset($m->d['a']) ? 'yes' : 'no';
"#,
        ["no"]
    };

    serialize_unserialize_object => {
        r#"<?php
class S { public function __construct(public int $n) {} }
$s = serialize(new S(9));
$o = unserialize($s);
echo $o->n;
"#,
        ["9"]
    };

    wakeup_mutates_after_unserialize => {
        r#"<?php
class S {
    public int $n = 0;
    public function __wakeup(): void { $this->n = 5; }
}
$o = unserialize(serialize(new S()));
echo $o->n;
"#,
        ["5"]
    };

    sleep_returns_property_names => {
        r#"<?php
class S {
    public int $a = 1;
    private int $b = 2;
    public function __sleep(): array { return ['a']; }
}
$s = serialize(new S());
echo str_contains($s, 'a') && !str_contains($s, 'b') ? 'trim' : 'full';
"#,
        ["trim"]
    };

    compare_objects_with_spaceship => {
        r#"<?php
class P { public function __construct(public int $n) {} }
echo (new P(1) <=> new P(2));
"#,
        ["-1"]
    };

    enum_in_class_constant => {
        r#"<?php
enum Color { case Red; }
class Palette { public const C = Color::Red; }
echo Palette::C->name;
"#,
        ["Red"]
    };

    readonly_promoted_property_access => {
        r#"<?php
readonly class Point { public function __construct(public int $x) {} }
echo (new Point(3))->x;
"#,
        ["3"]
    };

    allow_dynamic_properties_attribute => {
        r#"<?php
#[\AllowDynamicProperties]
class Flex {}
$f = new Flex();
$f->extra = 'ok';
echo $f->extra;
"#,
        ["ok"]
    };

    override_attribute_on_method => {
        r#"<?php
class Base { public function run(): string { return 'b'; } }
class Child extends Base { #[\Override] public function run(): string { return 'c'; } }
echo (new Child())->run();
"#,
        ["c"]
    };

    late_static_binding_static_call => {
        r#"<?php
class A { public static function who(): string { return static::class; } }
class B extends A {}
echo B::who();
"#,
        ["B"]
    };

    constant_on_class_via_const_keyword => {
        r#"<?php
class C { public const N = 7; }
echo C::N;
"#,
        ["7"]
    };
}
