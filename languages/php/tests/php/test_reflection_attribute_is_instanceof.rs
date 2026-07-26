
crate::php_cases! {
    reflection_attribute_is_instanceof => {
        r#"<?php
interface BaseAttr {}

#[Attribute]
class ConcreteAttr implements BaseAttr {}

#[ConcreteAttr]
class Subject {}

$rc = new ReflectionClass(Subject::class);

// Filter by exact class
$attrs1 = $rc->getAttributes(ConcreteAttr::class);
echo count($attrs1) . "|";

// Filter by interface using IS_INSTANCEOF
$attrs2 = $rc->getAttributes(BaseAttr::class, ReflectionAttribute::IS_INSTANCEOF);
echo count($attrs2) . "|";

// Filter by interface without IS_INSTANCEOF (should be 0)
$attrs3 = $rc->getAttributes(BaseAttr::class);
echo count($attrs3);
"#,
        ["1|1|0"]
    };
}
