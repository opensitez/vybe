//! How attributes behave across inheritance and traits.
//!
//! PHP does NOT inherit class-level attributes — Doctrine's mapped-superclass
//! support exists precisely because frameworks must walk `getParentClass()`
//! themselves. Method and property attributes behave differently again: they
//! belong to the *declaration*, so they stay visible through a child class.
//!
//! Expected values generated from PHP 8.4.11.

crate::php_cases! {
    class_attributes_are_not_inherited_by_a_child_class => {
        r#"<?php
#[Attribute]
class Entity {
    public function __construct(public string $table) {}
}
#[Entity('base_rows')]
class BaseModel {}
class Child extends BaseModel {}
echo count((new ReflectionClass(Child::class))->getAttributes(Entity::class));
"#,
        ["0"]
    };

    interface_attributes_are_not_inherited_by_implementor => {
        r#"<?php
#[Attribute]
class Contract {}
#[Contract]
interface Payable {}
class Invoice implements Payable {}
echo count((new ReflectionClass(Invoice::class))->getAttributes(Contract::class));
"#,
        ["0"]
    };

    parent_class_walk_collects_mapped_superclass_columns => {
        r#"<?php
#[Attribute]
class Column {
    public function __construct(public string $type) {}
}
class Timestamps {
    #[Column('datetime')]
    public $createdAt;
}
class Post extends Timestamps {
    #[Column('string')]
    public $title;
}
$names = [];
for ($rc = new ReflectionClass(Post::class); $rc; $rc = $rc->getParentClass()) {
    foreach ($rc->getProperties() as $p) {
        foreach ($p->getAttributes(Column::class) as $a) {
            $names[] = $rc->getName() . '.' . $p->getName() . ':' . $a->newInstance()->type;
        }
    }
}
echo implode('|', $names);
"#,
        ["Post.title:string|Post.createdAt:datetime|Timestamps.createdAt:datetime"]
    };

    inherited_method_attributes_are_visible_through_the_child => {
        r#"<?php
#[Attribute]
class Audit {
    public function __construct(public string $tag) {}
}
class Base {
    #[Audit('base-run')]
    public function run() {}
}
class Sub extends Base {}
$rm = new ReflectionMethod(Sub::class, 'run');
echo $rm->getAttributes(Audit::class)[0]->newInstance()->tag;
"#,
        ["base-run"]
    };

    overriding_method_replaces_the_parent_attributes => {
        r#"<?php
#[Attribute]
class Audit {
    public function __construct(public string $tag) {}
}
class Base {
    #[Audit('base')]
    public function run() {}
}
class Sub extends Base {
    #[Audit('sub')]
    public function run() {}
}
echo (new ReflectionMethod(Sub::class, 'run'))->getAttributes(Audit::class)[0]->newInstance()->tag;
"#,
        ["sub"]
    };

    trait_property_attributes_are_visible_on_the_using_class => {
        r#"<?php
#[Attribute]
class Column {
    public function __construct(public string $type) {}
}
trait HasId {
    #[Column('integer')]
    public $id;
}
class Thing {
    use HasId;
    #[Column('string')]
    public $name;
}
$out = [];
foreach ((new ReflectionClass(Thing::class))->getProperties() as $p) {
    foreach ($p->getAttributes(Column::class) as $a) {
        $out[] = $p->getName() . ':' . $a->newInstance()->type;
    }
}
echo implode('|', $out);
"#,
        ["name:string|id:integer"]
    };

    trait_method_attributes_are_visible_on_the_using_class => {
        r#"<?php
#[Attribute]
class Hook {
    public function __construct(public string $when) {}
}
trait Boots {
    #[Hook('boot')]
    public function boot() {}
}
class App {
    use Boots;
}
echo (new ReflectionMethod(App::class, 'boot'))->getAttributes(Hook::class)[0]->newInstance()->when;
"#,
        ["boot"]
    };
}
