//! The argument shapes framework attributes actually carry — enum cases,
//! class constants, nested arrays, mixed positional/named — and the split
//! between `getArguments()` (what was literally written) and `newInstance()`
//! (constructor defaults applied).
//!
//! Expected values generated from PHP 8.4.11.

crate::php_cases! {
    enum_case_argument_survives_new_instance => {
        r#"<?php
enum ColumnType: string {
    case Str = 'string';
    case Int = 'integer';
}
#[Attribute]
class Column {
    public function __construct(public ColumnType $type) {}
}
class Row {
    #[Column(ColumnType::Str)]
    public $name;
}
$a = (new ReflectionProperty(Row::class, 'name'))->getAttributes(Column::class)[0];
echo $a->newInstance()->type->value;
"#,
        ["string"]
    };

    class_constant_argument_is_resolved => {
        r#"<?php
class Limits {
    const MAX = 100;
}
#[Attribute]
class Cap {
    public function __construct(public int $n) {}
}
#[Cap(Limits::MAX)]
class Bucket {}
echo (new ReflectionClass(Bucket::class))->getAttributes(Cap::class)[0]->newInstance()->n;
"#,
        ["100"]
    };

    nested_array_argument_keeps_keys_and_depth => {
        r#"<?php
#[Attribute]
class Cfg {
    public function __construct(public array $opts) {}
}
#[Cfg(['db' => ['host' => 'localhost', 'port' => 5432], 'debug' => true])]
class Settings {}
$o = (new ReflectionClass(Settings::class))->getAttributes(Cfg::class)[0]->newInstance()->opts;
echo $o['db']['host'] . ':' . $o['db']['port'] . ':' . ($o['debug'] ? 'on' : 'off');
"#,
        ["localhost:5432:on"]
    };

    two_distinct_attribute_kinds_with_arguments_read_separately => {
        r#"<?php
#[Attribute]
class Route {
    public function __construct(public string $path) {}
}
#[Attribute]
class Auth {
    public function __construct(public string $role) {}
}
class Admin {
    #[Route('/admin')]
    #[Auth('superuser')]
    public function panel() {}
}
$rm = new ReflectionMethod(Admin::class, 'panel');
echo $rm->getAttributes(Route::class)[0]->newInstance()->path
   . '+' . $rm->getAttributes(Auth::class)[0]->newInstance()->role;
"#,
        ["/admin+superuser"]
    };

    omitted_default_is_absent_from_get_arguments_but_set_on_instance => {
        r#"<?php
#[Attribute]
class Opt {
    public function __construct(public int $a, public int $b = 9) {}
}
#[Opt(1)]
class One {}
$attr = (new ReflectionClass(One::class))->getAttributes(Opt::class)[0];
echo count($attr->getArguments()) . ':' . $attr->newInstance()->b;
"#,
        ["1:9"]
    };

    named_argument_is_keyed_by_name_in_get_arguments => {
        r#"<?php
#[Attribute]
class Mix {
    public function __construct(public int $a, public string $b) {}
}
#[Mix(5, b: 'x')]
class M {}
$args = (new ReflectionClass(M::class))->getAttributes(Mix::class)[0]->getArguments();
echo $args[0] . ',' . $args['b'];
"#,
        ["5,x"]
    };

    bool_null_and_float_arguments_round_trip => {
        r#"<?php
#[Attribute]
class Types {
    public function __construct(public bool $flag, public ?string $none, public float $f) {}
}
#[Types(true, null, 1.5)]
class T {}
$i = (new ReflectionClass(T::class))->getAttributes(Types::class)[0]->newInstance();
echo var_export($i->flag, true) . ',' . var_export($i->none, true) . ',' . $i->f;
"#,
        ["true,NULL,1.5"]
    };
}
