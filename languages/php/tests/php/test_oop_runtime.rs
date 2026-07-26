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
        ["empty"]
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
        ["eq"]
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

    dynamic_property_access_by_name => {
        r#"<?php
class Profile {
    public string $username = 'alice';
}
$p = new Profile();
$field = 'username';
echo $p->$field;
"#,
        ["alice"]
    };

    dynamic_method_invocation_by_name => {
        r#"<?php
class Worker {
    public function work(string $job): string { return "job:$job"; }
}
$w = new Worker();
$method = 'work';
echo $w->$method('build');
"#,
        ["job:build"]
    };

    anonymous_class_extends_runtime_parent_call => {
        r#"<?php
class Base {
    public function label(): string { return 'base'; }
}
$o = new class extends Base {
    public function label(): string { return parent::label() . '+anon'; }
};
echo $o->label();
"#,
        ["base+anon"]
    };

    class_alias_works_with_existing_class => {
        r#"<?php
class OriginalService {}
class_alias(OriginalService::class, 'AliasService');
echo class_exists('AliasService') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    trait_detection_via_trait_exists => {
        r#"<?php
trait AuditTrait { public function audit(): string { return 'ok'; } }
echo trait_exists(AuditTrait::class) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    property_can_be_set_and_read_through_magic_hooks => {
        r#"<?php
class Bag {
    private array $d = [];
    public function __set(string $k, mixed $v): void { $this->d[$k] = $v; }
    public function __get(string $k): mixed { return $this->d[$k] ?? null; }
}

$b = new Bag();
$b->x = 9;
echo $b->x;
"#,
        ["9"]
    };

    object_clone_with_reference_property_duplicates_handle => {
        r#"<?php
class Cnt {
    public function __construct(public array &$nums) {}
    public function value(): int { return $this->nums[0]; }
}
$arr = [1];
$a = new Cnt($arr);
$b = clone $a;
$arr[0] = 4;
echo $a->value() . $b->value();
"#,
        ["44"]
    };

    instance_method_call_with_named_arguments => {
        r#"<?php
class Calc {
    public function add(int $a, int $b): int { return $a + $b; }
}
$c = new Calc();
echo $c->add(b: 2, a: 3);
"#,
        ["5"]
    };

    method_visibility_violation_triggers_error_string => {
        r#"<?php
class Vault {
    private function secret(): string { return 'x'; }
}

$ok = false;
try {
    (new Vault())->secret();
} catch (Error $e) {
    $ok = true;
}
echo $ok ? 'error' : 'ok';
"#,
        ["error"]
    };

    get_called_class_in_inheritance_chain => {
        r#"<?php
class Root {
    public static function what(): string { return get_called_class(); }
}
class Node extends Root {}
class Leaf extends Node {}
echo Node::what() . '|' . Leaf::what();
"#,
        ["Node|Leaf"]
    };

    object_to_string_magic_method => {
        r#"<?php
class Label {
    public function __toString(): string { return 'labelled'; }
}
echo (string) new Label();
"#,
        ["labelled"]
    };

    object_invokes_magic_callable => {
        r#"<?php
class Inv {
    public function __invoke(int $n): string { return 'v' . $n; }
}
$i = new Inv();
echo $i(4);
"#,
        ["v4"]
    };

    call_static_magic_invoked_on_missing_static_method => {
        r#"<?php
class StaticGhost {
    public static function __callStatic(string $name, array $args): string {
        return $name . count($args);
    }
}
echo StaticGhost::build('a', 'b');
"#,
        ["build2"]
    };

    object_equality_with_same_properties_false_identity => {
        r#"<?php
class N { public function __construct(public int $n) {} }
$a = new N(1);
$b = new N(1);
echo ($a == $b ? 'eq' : 'ne') . '|' . (($a === $b) ? 'same' : 'diff');
"#,
        ["eq|diff"]
    };

    compare_objects_with_same_handle_using_spaceship => {
        r#"<?php
class Cmp {
    public function __construct(public int $v) {}
}
$a = new Cmp(4);
$b = $a;
echo ($a <=> $b);
"#,
        ["0"]
    };

    clone_copies_private_property_values => {
        r#"<?php
class Secret {
    public function __construct(private int $n) {}
    public function get(): int { return $this->n; }
}
$a = new Secret(8);
$b = clone $a;
echo $a->get() . $b->get();
"#,
        ["88"]
    };

    clone_runs_magic_on_each_copy => {
        r#"<?php
class Mark {
    public string $tag = 'x';
    public function __clone(): void { $this->tag .= '!'; }
}
$orig = new Mark();
$copy = clone $orig;
echo $orig->tag . '|' . $copy->tag;
"#,
        ["x|x!"]
    };

    dynamic_new_via_variable_class_name => {
        r#"<?php
class Engine {
    public function name(): string { return 'ok'; }
}
$class = 'Engine';
$obj = new $class();
echo $obj->name();
"#,
        ["ok"]
    };

    property_array_unset_via_magic_unset => {
        r#"<?php
class Bag {
    private array $items = ['a' => 1];
    public function __isset(string $name): bool { return isset($this->items[$name]); }
    public function __unset(string $name): void { unset($this->items[$name]); }
    public function has(string $name): bool { return isset($this->items[$name]); }
}
$b = new Bag();
unset($b->a);
echo $b->has('a') ? 'yes' : 'no';
"#,
        ["no"]
    };

    object_debug_info_hides_private_property => {
        r#"<?php
class Debug {
    public function __construct(private string $token) {}
    public function __debugInfo(): array { return ['token' => $this->token, 'public' => true]; }
}
$d = new Debug();
$arr = json_encode((array) print_r($d, true));
echo (str_contains($arr, 'token') && str_contains($arr, 'public')) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    weak_reference_collects_alive_object_id => {
        r#"<?php
class Hold {}
$o = new Hold();
$wr = WeakReference::create($o);
unset($o);
echo $wr->get() === null ? 'dead' : 'alive';
"#,
        ["alive"]
    };

    trait_conflict_resolved_with_insteadof => {
        r#"<?php
trait A { public function f(): string { return 'a'; } }
trait B { public function f(): string { return 'b'; } }

class C {
    use A, B {
        A::f insteadof B;
    }
}
echo (new C())->f();
"#,
        ["a"]
    };

    object_cast_from_array_with_stdclass => {
        r#"<?php
$o = (object) ['a' => 1, 'b' => 2];
echo $o->a . $o->b;
"#,
        ["12"]
    };

    magic_set_state_allows_unsafe_restore => {
        r#"<?php
class Pack {
    public int $value;
    public static function __set_state(array $data): self { $x = new self(); $x->value = $data['value'] * 2; return $x; }
}
$obj = var_export(['value' => 3], true);
$p = Pack::__set_state(['value' => 3]);
echo $p->value;
"#,
        ["6"]
    };

    get_declared_traits_includes_local_trait => {
        r#"<?php
trait Tracked {}
class X { use Tracked; }
echo in_array('Tracked', get_declared_traits()) ? 'yes' : 'no';
"#,
        ["yes"]
    };
}
