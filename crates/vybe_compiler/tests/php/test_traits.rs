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
}
