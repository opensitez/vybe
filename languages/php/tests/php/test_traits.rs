//! `trait` use, aliases, insteadof, and horizontal composition.

crate::php_cases! {
    trait_method_used_by_class => {
        r#"<?php
trait Speak {
    public function speak(): string { return 'hi'; }
}
class Bot {
    use Speak;
}
echo (new Bot())->speak();
"#,
        ["hi"]
    };

    trait_property_shared_via_use => {
        r#"<?php
trait Counter {
    public int $n = 0;
}
class Box {
    use Counter;
}
$b = new Box();
$b->n = 7;
echo $b->n;
"#,
        ["7"]
    };

    trait_insteadof_selects_preferred_method => {
        r#"<?php
trait A { public function run(): string { return 'a'; } }
trait B { public function run(): string { return 'b'; } }
class Worker {
    use A, B { A::run insteadof B; }
}
echo (new Worker())->run();
"#,
        ["a"]
    };

    trait_alias_renames_method => {
        r#"<?php
trait Base {
    public function work(): string { return 'done'; }
}
class Job {
    use Base { work as perform; }
}
echo (new Job())->perform();
"#,
        ["done"]
    };

    trait_alias_with_visibility_change => {
        r#"<?php
trait Secret {
    private function token(): string { return 'tok'; }
}
class Api {
    use Secret { token as public reveal; }
}
echo (new Api())->reveal();
"#,
        ["tok"]
    };

    trait_nested_use_composes_behaviors => {
        r#"<?php
trait One { public function a(): int { return 1; } }
trait Two { use One; public function b(): int { return 2; } }
class Both { use Two; }
$o = new Both();
echo $o->a() + $o->b();
"#,
        ["3"]
    };

    trait_abstract_method_implemented_by_class => {
        r#"<?php
trait Template {
    abstract public function body(): string;
    public function render(): string { return '<' . $this->body() . '>'; }
}
class Page {
    use Template;
    public function body(): string { return 'main'; }
}
echo (new Page())->render();
"#,
        ["<main>"]
    };

    trait_static_method_callable => {
        r#"<?php
trait Factory {
    public static function make(): string { return 'new'; }
}
class Item { use Factory; }
echo Item::make();
"#,
        ["new"]
    };

    trait_constants_accessible_from_class => {
        r#"<?php
trait Limits { public const MAX = 10; }
class Batch { use Limits; }
echo Batch::MAX;
"#,
        ["10"]
    };

    trait_precedence_parent_class_wins_without_insteadof => {
        r#"<?php
class Base { public function id(): string { return 'base'; } }
trait T { public function id(): string { return 'trait'; } }
class Child extends Base { use T; }
echo (new Child())->id();
"#,
        ["trait"]
    };

    trait_multiple_uses_same_trait_ok => {
        r#"<?php
trait Log { public function log(): string { return 'l'; } }
class App { use Log; }
echo (new App())->log();
"#,
        ["l"]
    };

    trait_method_calls_other_trait_method => {
        r#"<?php
trait A { public function x(): int { return 1; } }
trait B {
    use A;
    public function y(): int { return $this->x() + 1; }
}
class C { use B; }
echo (new C())->y();
"#,
        ["2"]
    };

    trait_private_helper_exposed_via_public_alias => {
        r#"<?php
trait Core {
    private function token(): string { return 'tk'; }
}
class Service {
    use Core { token as public get_token; }
}
echo (new Service())->get_token();
"#,
        ["tk"]
    };

    trait_insteadof_prevents_class_parent_collision => {
        r#"<?php
trait A { public function label(): string { return 'a'; } }
trait B { public function label(): string { return 'b'; } }
class App {
    use A, B { A::label insteadof B; }
}
echo (new App())->label();
"#,
        ["a"]
    };

    trait_chained_aliases_in_single_class => {
        r#"<?php
trait T {
    public function work(): string { return 'w'; }
}
class W {
    use T {
        work as execute;
    }
}
$w = new W();
echo $w->work() . $w->execute();
"#,
        ["ww"]
    };

    trait_mutable_shared_state => {
        r#"<?php
trait Counter {
    public int $count = 0;
    public function inc(): int { return ++$this->count; }
}
class Worker {
    use Counter;
}
$a = new Worker();
$b = new Worker();
echo $a->inc() . $b->inc();
"#,
        ["11"]
    };

    trait_property_defaults_and_override => {
        r#"<?php
trait Defaults {
    public function base(): string { return 'd'; }
}
class ParentThing {
    public string $v = 'parent';
}
class ChildThing extends ParentThing {
    use Defaults;
    public string $v = 'child';
}
echo (new ChildThing())->v . (new ChildThing())->base();
"#,
        ["childd"]
    };

    trait_stateful_method_visibility => {
        r#"<?php
trait Hidden {
    public function open(): string { return 'open'; }
    private function secret(): string { return 'secret'; }
    public function reveal(): string { return $this->secret(); }
}
class Door { use Hidden; }
echo (new Door())->open() . '|' . (new Door())->reveal();
"#,
        ["open|secret"]
    };

    trait_instantiate_static_context => {
        r#"<?php
trait Maker {
    public static function create(int $value): static {
        return new static($value);
    }
}
class Item {
    use Maker;
    public function __construct(private int $v) {}
    public function value(): int { return $this->v; }
}
echo Item::create(7)->value();
"#,
        ["7"]
    };

    trait_overrides_from_parent_then_restores_parent => {
        r#"<?php
class Base {
    public function label(): string { return 'base'; }
}
trait T {
    public function label(): string { return 'trait'; }
}
class Child extends Base {
    use T {
        T::label as parentLabel;
    }
    public function local(): string { return $this->parentLabel(); }
}
$obj = new Child();
echo $obj->label() . '|' . $obj->local();
"#,
        ["trait|base"]
    };

    trait_renamed_static_method => {
        r#"<?php
trait Printer {
    public static function asText(): string { return 'text'; }
}
class Doc {
    use Printer { asText as public format; }
}
echo Doc::format();
"#,
        ["text"]
    };

    trait_static_and_instance_coexist => {
        r#"<?php
trait Marker {
    public function ping(): string { return 'instance'; }
    public static function pingStatic(): string { return 'static'; }
}
class Endpoint {
    use Marker;
}
echo (new Endpoint())->ping() . '|' . Endpoint::pingStatic();
"#,
        ["instance|static"]
    };

    trait_nested_alias_chain => {
        r#"<?php
trait BaseFlow {
    public function phase(): string { return 'base'; }
}
trait Flow {
    use BaseFlow { phase as basePhase; }
    public function phase(): string { return 'flow'; }
}
class Pipeline {
    use Flow;
    public function run(): string { return $this->phase() . '|' . $this->basePhase(); }
}
echo (new Pipeline())->run();
"#,
        ["flow|base"]
    };

    trait_intersection_with_interfaces_runtime => {
        r#"<?php
interface Loggable { public function log(): string; }
trait Logger {
    public function log(): string { return 'log'; }
}
class Service implements Loggable {
    use Logger;
}
echo (new Service())->log();
"#,
        ["log"]
    };

    trait_multiple_inheritance_of_traits_order => {
        r#"<?php
trait A { public function stamp(): string { return 'a'; } }
trait B { public function stamp(): string { return 'b'; } }
trait C { use A, B { A::stamp insteadof B; B::stamp as fromB; } }
class Recorder { use C; }
$r = new Recorder();
echo $r->stamp() . '|' . $r->fromB();
"#,
        ["a|b"]
    };

    trait_property_initializer_shares_per_instance_state => {
        r#"<?php
trait BoxState {
    public int $count = 0;
    public function inc(): int { return ++$this->count; }
}
class Bucket {
    use BoxState;
}
$a = new Bucket();
$b = new Bucket();
echo $a->inc() . '|' . $b->inc() . '|' . $a->inc();
"#,
        ["1|1|2"]
    };
}
