use super::helpers::run_prints;

crate::php_cases! {
    reflection_class_get_attributes_empty => {
        r#"<?php
class NoAttrClass {}
$rc = new ReflectionClass(NoAttrClass::class);
echo count($rc->getAttributes());
"#,
        ["0"]
    };

    reflection_class_get_attributes_multiple => {
        r#"<?php
#[Attribute]
class MyAttr1 {}

#[Attribute]
class MyAttr2 {}

#[MyAttr1]
#[MyAttr2]
class TargetClass {}

$rc = new ReflectionClass(TargetClass::class);
$attrs = $rc->getAttributes();
echo count($attrs) . "|";
echo $attrs[0]->getName() . "|" . $attrs[1]->getName();
"#,
        ["2|MyAttr1|MyAttr2"]
    };

    reflection_class_get_attributes_with_arguments => {
        r#"<?php
#[Attribute]
class ArgAttr {
    public function __construct(public string $val) {}
}

#[ArgAttr('hello')]
class TargetClassArgs {}

$rc = new ReflectionClass(TargetClassArgs::class);
$attrs = $rc->getAttributes();
$args = $attrs[0]->getArguments();
echo $args[0];
"#,
        ["hello"]
    };
}
