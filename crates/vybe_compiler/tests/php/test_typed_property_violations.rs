//! Typed / readonly property access violations — distinct runtime Error paths.

crate::php_cases! {
    read_uninitialized_typed_instance_property => {
        r#"<?php
class Box { public int $size; }
$o = new Box();
try { echo $o->size; }
catch (Error $e) { echo 'uninit'; }
"#,
        ["uninit"]
    };

    write_uninitialized_typed_instance_property => {
        r#"<?php
class Box { public string $label; }
$o = new Box();
$o->label = 'ok';
echo $o->label;
"#,
        ["ok"]
    };

    read_typed_property_after_write => {
        r#"<?php
class Pair { public int $a; public int $b; }
$p = new Pair();
$p->a = 1;
$p->b = 2;
echo $p->a + $p->b;
"#,
        ["3"]
    };

    readonly_property_second_assignment_fails => {
        r#"<?php
readonly class Token { public function __construct(public string $value) {} }
$t = new Token('abc');
try { $t->value = 'xyz'; echo 'mutated'; }
catch (Error $e) { echo 'readonly'; }
"#,
        ["readonly"]
    };

    readonly_clone_creates_independent_copy => {
        r#"<?php
readonly class Point { public function __construct(public int $x) {} }
$p1 = new Point(1);
$p2 = clone $p1;
echo $p2->x;
"#,
        ["1"]
    };

    static_typed_property_uninitialized_read => {
        r#"<?php
class Cache { public static string $key; }
try { echo Cache::$key; }
catch (Error $e) { echo 'static-uninit'; }
"#,
        ["static-uninit"]
    };

    static_typed_property_write_then_read => {
        r#"<?php
class Cache { public static int $hits; }
Cache::$hits = 3;
echo Cache::$hits;
"#,
        ["3"]
    };

    private_typed_property_via_getter_only => {
        r#"<?php
class Wallet { private int $balance = 0; public function credit(int $n): void { $this->balance += $n; } public function total(): int { return $this->balance; } }
$w = new Wallet();
$w->credit(5);
echo $w->total();
"#,
        ["5"]
    };

    protected_typed_child_access => {
        r#"<?php
class Base { protected int $n = 7; }
class Child extends Base { public function expose(): int { return $this->n; } }
echo (new Child())->expose();
"#,
        ["7"]
    };

    nullable_typed_defaults_to_null_without_error => {
        r#"<?php
class Maybe { public ?string $note; }
$m = new Maybe();
echo $m->note === null ? 'null' : 'set';
"#,
        ["null"]
    };

    nullable_typed_assign_string => {
        r#"<?php
class Maybe { public ?string $note; }
$m = new Maybe();
$m->note = 'hi';
echo $m->note;
"#,
        ["hi"]
    };

    union_typed_property_accepts_int_branch => {
        r#"<?php
class Flex { public int|string $id; }
$f = new Flex();
$f->id = 42;
echo $f->id;
"#,
        ["42"]
    };

    union_typed_property_accepts_string_branch => {
        r#"<?php
class Flex { public int|string $id; }
$f = new Flex();
$f->id = 'uuid';
echo $f->id;
"#,
        ["uuid"]
    };

    intersection_typed_property_requires_both => {
        r#"<?php
interface A { public function a(): void; }
interface B { public function b(): void; }
class Both implements A, B { public function a(): void {} public function b(): void {} }
class Holder { public A&B $item; }
$h = new Holder();
$h->item = new Both();
echo $h->item instanceof Both ? 'both' : 'no';
"#,
        ["both"]
    };

    promoted_typed_constructor_property_read => {
        r#"<?php
class User { public function __construct(public string $name) {} }
$u = new User('ana');
echo $u->name;
"#,
        ["ana"]
    };

    promoted_readonly_cannot_reassign => {
        r#"<?php
class User { public function __construct(public readonly int $id) {} }
$u = new User(1);
try { $u->id = 2; echo 'changed'; }
catch (Error $e) { echo 'blocked'; }
"#,
        ["blocked"]
    };

    dynamic_property_on_typed_class_without_allow_dynamic => {
        r#"<?php
class Strict { public int $n = 1; }
$s = new Strict();
try { $s->extra = 2; echo 'added'; }
catch (Error $e) { echo 'dynamic'; }
"#,
        ["dynamic"]
    };

    std_class_allows_dynamic_properties => {
        r#"<?php
$o = new stdClass();
$o->x = 9;
echo $o->x;
"#,
        ["9"]
    };

    typed_property_with_default_skips_uninit_error => {
        r#"<?php
class Counter { public int $n = 0; }
echo (new Counter())->n;
"#,
        ["0"]
    };

    typed_array_property_default_empty => {
        r#"<?php
class Bag { public array $items = []; }
$b = new Bag();
echo count($b->items);
"#,
        ["0"]
    };

    typed_array_property_push_after_init => {
        r#"<?php
class Bag { public array $items; }
$b = new Bag();
$b->items = [1];
$b->items[] = 2;
echo implode(',', $b->items);
"#,
        ["1,2"]
    };

    readonly_class_all_properties_readonly => {
        r#"<?php
readonly class Config { public function __construct(public int $port, public string $host) {} }
$c = new Config(8080, 'localhost');
echo $c->port . ':' . $c->host;
"#,
        ["8080:localhost"]
    };

    readonly_extends_readonly_child_assign_fails => {
        r#"<?php
readonly class Base { public function __construct(public int $v) {} }
readonly class Child extends Base {}
$c = new Child(5);
try { $c->v = 6; echo 'ok'; }
catch (Error $e) { echo 'fail'; }
"#,
        ["fail"]
    };

    uninitialized_property_in_nested_object_graph => {
        r#"<?php
class Inner { public int $x; }
class Outer { public Inner $inner; }
$o = new Outer();
$o->inner = new Inner();
try { echo $o->inner->x; }
catch (Error $e) { echo 'nested'; }
"#,
        ["nested"]
    };

    isset_on_uninitialized_typed_property_is_false => {
        r#"<?php
class Slot { public int $n; }
$s = new Slot();
echo isset($s->n) ? 'set' : 'unset';
"#,
        ["unset"]
    };

    property_exists_on_uninitialized_typed_is_true => {
        r#"<?php
class Slot { public int $n; }
$s = new Slot();
echo property_exists($s, 'n') ? 'exists' : 'missing';
"#,
        ["exists"]
    };

    unset_uninitialized_typed_then_read_still_fails => {
        r#"<?php
class Slot { public int $n; }
$s = new Slot();
unset($s->n);
try { echo $s->n; }
catch (Error $e) { echo 'still'; }
"#,
        ["still"]
    };

    typed_property_in_abstract_class_concrete_init => {
        r#"<?php
abstract class Shape { public int $sides; }
class Tri extends Shape {}
$t = new Tri();
$t->sides = 3;
echo $t->sides;
"#,
        ["3"]
    };

    interface_typed_property_on_implementation => {
        r#"<?php
interface HasId { public int $id; }
class Entity implements HasId { public int $id; }
$e = new Entity();
$e->id = 99;
echo $e->id;
"#,
        ["99"]
    };

    enum_backed_property_not_confused_with_typed_field => {
        r#"<?php
enum Status: string { case On = 'on'; }
class Machine { public Status $state; }
$m = new Machine();
$m->state = Status::On;
echo $m->state->value;
"#,
        ["on"]
    };

    weak_typed_reference_property_requires_init => {
        r#"<?php
class Node { public WeakReference $ref; }
$n = new Node();
try { echo $n->ref; }
catch (Error $e) { echo 'weak'; }
"#,
        ["weak"]
    };

    typed_property_hooked_accessor_read => {
        r#"<?php
class Meter {
    private int $v = 0;
    public int $reading {
        get => $this->v;
        set => $this->v = $value;
    }
}
$m = new Meter();
$m->reading = 4;
echo $m->reading;
"#,
        ["4"]
    };

    readonly_hooked_property_cannot_set => {
        r#"<?php
class RO {
    public function __construct(public readonly int $id) {}
}
$r = new RO(1);
try { $r->id = 2; }
catch (Error $e) { echo 'ro'; }
"#,
        ["ro"]
    };

    serializable_typed_property_after_unserialize => {
        r#"<?php
class Data { public int $n; }
$d = new Data();
$d->n = 8;
$copy = unserialize(serialize($d));
echo $copy->n;
"#,
        ["8"]
    };

    clone_copies_initialized_typed_values => {
        r#"<?php
class Cell { public int $v; }
$a = new Cell();
$a->v = 2;
$b = clone $a;
$b->v = 5;
echo $a->v . ':' . $b->v;
"#,
        ["2:5"]
    };

    typed_parameter_property_promotion_with_default_nullable => {
        r#"<?php
class Opt { public function __construct(public ?float $rate = null) {} }
$o = new Opt();
echo $o->rate === null ? 'none' : 'set';
"#,
        ["none"]
    };

    accessing_typed_property_on_uninitialized_this_in_method => {
        r#"<?php
class Late { public int $n; public function peek(): void { try { echo $this->n; } catch (Error $e) { echo 'late'; } } }
(new Late())->peek();
"#,
        ["late"]
    };

    assigning_null_to_non_nullable_typed_property_type_error => {
        r#"<?php
class Strict { public string $name; }
$s = new Strict();
try { $s->name = null; echo 'ok'; }
catch (TypeError $e) { echo 'type'; }
"#,
        ["type"]
    };

    assigning_wrong_scalar_type_to_typed_property => {
        r#"<?php
class Strict { public int $count; }
$s = new Strict();
try { $s->count = 'x'; echo 'ok'; }
catch (TypeError $e) { echo 'scalar'; }
"#,
        ["scalar"]
    };

    assigning_array_to_scalar_typed_property_fails => {
        r#"<?php
class Strict { public float $rate; }
$s = new Strict();
try { $s->rate = []; echo 'ok'; }
catch (TypeError $e) { echo 'array'; }
"#,
        ["array"]
    };

    bool_typed_property_rejects_string_one => {
        r#"<?php
class Flag { public bool $on; }
$f = new Flag();
try { $f->on = '1'; echo 'ok'; }
catch (TypeError $e) { echo 'bool'; }
"#,
        ["bool"]
    };

    callable_typed_property_accepts_closure => {
        r#"<?php
class Runner { public callable $fn; }
$r = new Runner();
$r->fn = fn(int $n) => $n + 1;
echo ($r->fn)(4);
"#,
        ["5"]
    };

    object_typed_property_requires_instance => {
        r#"<?php
class Holder { public stdClass $obj; }
$h = new Holder();
try { $h->obj = 1; echo 'ok'; }
catch (TypeError $e) { echo 'object'; }
"#,
        ["object"]
    };

    iterable_typed_property_accepts_array => {
        r#"<?php
class Coll { public iterable $items; }
$c = new Coll();
$c->items = [1, 2];
echo is_array($c->items) ? 'iter' : 'no';
"#,
        ["iter"]
    };

    mixed_typed_property_accepts_any_value => {
        r#"<?php
class Any { public mixed $slot; }
$a = new Any();
$a->slot = new ArrayObject([1]);
echo $a->slot instanceof ArrayObject ? 'mixed' : 'no';
"#,
        ["mixed"]
    };

    false_typed_property_only_accepts_false => {
        r#"<?php
class Off { public false $state; }
$o = new Off();
$o->state = false;
echo $o->state === false ? 'false' : 'other';
"#,
        ["false"]
    };

    null_typed_property_only_accepts_null => {
        r#"<?php
class Empty { public null $value; }
$e = new Empty();
$e->value = null;
echo $e->value === null ? 'null' : 'other';
"#,
        ["null"]
    };

    never_returning_method_stops_normal_flow => {
        r#"<?php
class Stop { public int $n = 1; public function halt(): never { throw new RuntimeException('stop'); } }
$s = new Stop();
try { $s->halt(); } catch (RuntimeException $e) { echo 'never'; }
"#,
        ["never"]
    };

    typed_static_property_shared_across_instances => {
        r#"<?php
class Shared { public static int $total = 0; }
Shared::$total = 4;
echo Shared::$total;
"#,
        ["4"]
    };

    parent_private_not_visible_to_child_typed_access => {
        r#"<?php
class ParentBox { private int $secret = 1; }
class ChildBox extends ParentBox { public function leak(): void { try { echo $this->secret; } catch (Error $e) { echo 'private'; } } }
(new ChildBox())->leak();
"#,
        ["private"]
    };

    trait_typed_property_on_using_class => {
        r#"<?php
trait Counter { public int $n; }
class App { use Counter; }
$a = new App();
$a->n = 6;
echo $a->n;
"#,
        ["6"]
    };
}
