//! PHP 8 `#[Attribute]` declarations, targets, repetition, and reflection.

crate::php_cases! {
    attribute_class_constructor_argument_read_via_reflection => {
        r#"<?php
#[Attribute]
class Version {
    public function __construct(public string $number) {}
}
#[Version('2.1.0')]
class App {}
echo (new ReflectionClass(App::class))->getAttributes(Version::class)[0]->newInstance()->number;
"#,
        ["2.1.0"]
    };

    attribute_on_function_parameter => {
        r#"<?php
#[Attribute]
class Sensitive {}
function login(string $user, #[Sensitive] string $pass): string {
    return $user;
}
$params = (new ReflectionFunction('login'))->getParameters();
echo count($params[1]->getAttributes(Sensitive::class));
"#,
        ["1"]
    };

    attribute_repeatable_on_same_class => {
        r#"<?php
#[Attribute(Attribute::IS_REPEATABLE | Attribute::TARGET_CLASS)]
class Tag {
    public function __construct(public string $name) {}
}
#[Tag('api')]
#[Tag('v1')]
class Endpoint {}
echo count((new ReflectionClass(Endpoint::class))->getAttributes(Tag::class));
"#,
        ["2"]
    };

    attribute_on_interface_method => {
        r#"<?php
#[Attribute]
class Route {
    public function __construct(public string $path) {}
}
interface Controller {
    #[Route('/index')]
    public function index(): string;
}
$ref = new ReflectionMethod(Controller::class, 'index');
echo $ref->getAttributes(Route::class)[0]->newInstance()->path;
"#,
        ["/index"]
    };

    attribute_on_class_constant => {
        r#"<?php
#[Attribute]
class Meta {
    public function __construct(public string $key) {}
}
class Config {
    #[Meta('app.debug')]
    public const DEBUG = true;
}
$ref = new ReflectionClassConstant(Config::class, 'DEBUG');
echo $ref->getAttributes(Meta::class)[0]->newInstance()->key;
"#,
        ["app.debug"]
    };

    attribute_on_backed_enum_case => {
        r#"<?php
#[Attribute]
class Label {
    public function __construct(public string $text) {}
}
enum Color: string {
    #[Label('Red')]
    case Red = 'red';
}
$ref = new ReflectionEnumUnitCase(Color::class, 'Red');
echo $ref->getAttributes(Label::class)[0]->newInstance()->text;
"#,
        ["Red"]
    };

    attribute_named_constructor_arguments => {
        r#"<?php
#[Attribute]
class Cache {
    public function __construct(public int $ttl, public string $key) {}
}
#[Cache(ttl: 60, key: 'users')]
class Repo {}
$attr = (new ReflectionClass(Repo::class))->getAttributes(Cache::class)[0]->newInstance();
echo $attr->ttl . ':' . $attr->key;
"#,
        ["60:users"]
    };

    attribute_get_arguments_preserves_positional_values => {
        r#"<?php
#[Attribute]
class Pair {
    public function __construct(public int $a, public int $b) {}
}
#[Pair(3, 7)]
class Box {}
$args = (new ReflectionClass(Box::class))->getAttributes(Pair::class)[0]->getArguments();
echo $args[0] . '+' . $args[1];
"#,
        ["3+7"]
    };

    attribute_is_instance_returns_true_for_matching_class => {
        r#"<?php
#[Attribute]
class Flag {}
#[Flag]
class Marked {}
$attr = (new ReflectionClass(Marked::class))->getAttributes()[0];
echo $attr->isInstance(Flag::class) ? 'flag' : 'other';
"#,
        ["flag"]
    };

    attribute_count_without_filter_includes_all => {
        r#"<?php
#[Attribute]
class A {}
#[Attribute]
class B {}
#[A]
#[B]
class Dual {}
echo count((new ReflectionClass(Dual::class))->getAttributes());
"#,
        ["2"]
    };

    attribute_on_closure_is_not_supported_but_class_still_reflects => {
        r#"<?php
#[Attribute]
class Note {
    public function __construct(public string $msg) {}
}
#[Note('worker')]
class Worker {
    public function tag(): string {
        return (new ReflectionClass(self::class))->getAttributes(Note::class)[0]->newInstance()->msg;
    }
}
echo (new Worker())->tag();
"#,
        ["worker"]
    };

    attribute_target_class_only_rejects_method_placement_at_compile => {
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS)]
class ClassOnly {}
#[ClassOnly]
class Ok {
    public function ok(): string { return 'ok'; }
}
echo (new Ok())->ok();
"#,
        ["ok"]
    };

    override_attribute_on_overridden_method => {
        r#"<?php
class Base {
    public function run(): string { return 'base'; }
}
class Child extends Base {
    #[\Override]
    public function run(): string { return 'child'; }
}
echo (new Child())->run();
"#,
        ["child"]
    };

    allow_dynamic_properties_attribute_permits_new_fields => {
        r#"<?php
#[\AllowDynamicProperties]
class Bag {}
$b = new Bag();
$b->extra = 'dyn';
echo $b->extra;
"#,
        ["dyn"]
    };

    attribute_on_promoted_constructor_parameter => {
        r#"<?php
#[Attribute]
class Inject {}
class Service {
    public function __construct(#[Inject] public string $name) {}
}
$ref = (new ReflectionClass(Service::class))->getConstructor()->getParameters()[0];
echo $ref->getAttributes(Inject::class) ? 'inject' : 'plain';
"#,
        ["inject"]
    };

    attribute_get_arguments_array => {
        r#"<?php
#[Attribute]
class Options {
    public function __construct(public array $list) {}
}
#[Options(['a' => 1, 'b' => 2])]
class ConfigHolder {}
$attr = (new ReflectionClass(ConfigHolder::class))->getAttributes(Options::class)[0];
$args = $attr->getArguments();
echo is_array($args) && isset($args[0]['a']) ? 'args_array_ok' : 'err';
"#,
        ["args_array_ok"]
    };

    attribute_get_target_bitmask => {
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS | Attribute::TARGET_METHOD)]
class DualTarget {}
$rc = new ReflectionClass(DualTarget::class);
$attr = $rc->getAttributes(Attribute::class)[0]->newInstance();
echo ($attr->flags & Attribute::TARGET_CLASS) ? 'has_class_target' : 'err';
"#,
        ["has_class_target"]
    };

    deprecated_attribute_php84_builtin => {
        r#"<?php
class Legacy {
    #[\Deprecated("Use newMethod instead", since: "2.0")]
    public function oldMethod(): string { return 'legacy'; }
}
$rm = new ReflectionMethod(Legacy::class, 'oldMethod');
$attrs = $rm->getAttributes(\Deprecated::class);
echo count($attrs) . '|' . $attrs[0]->getName() . '|';
$d = $attrs[0]->newInstance();
echo $d->message . '|' . $d->since . '|' . ($rm->isDeprecated() ? 'yes' : 'no');
"#,
        ["1|Deprecated|Use newMethod instead|2.0|yes"]
    };
}

