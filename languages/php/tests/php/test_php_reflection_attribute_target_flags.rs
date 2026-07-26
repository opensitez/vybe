use super::helpers::run_prints;

#[test]
fn test_reflection_attribute_target_class() {
    assert_eq!(
        run_prints(
            r#"<?php
#[Attribute(Attribute::TARGET_CLASS)]
class CustomAttr {}

#[CustomAttr]
class TargetClass {}

$rc = new ReflectionClass(TargetClass::class);
$attrs = $rc->getAttributes(CustomAttr::class);
echo count($attrs) . ':' . $attrs[0]->getName(), "\n";
"#
        ),
        vec!["1:CustomAttr"]
    );
}

#[test]
fn test_reflection_attribute_instantiate() {
    assert_eq!(
        run_prints(
            r#"<?php
#[Attribute]
class ParamAttr {
    public function __construct(public string $label) {}
}

#[ParamAttr('demo_label')]
class Demo {}

$rc = new ReflectionClass(Demo::class);
$attr = $rc->getAttributes(ParamAttr::class)[0]->newInstance();
echo $attr->label, "\n";
"#
        ),
        vec!["demo_label"]
    );
}
