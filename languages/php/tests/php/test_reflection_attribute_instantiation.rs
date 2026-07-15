use super::helpers::run_prints;

crate::php_cases! {
    reflection_attribute_new_instance => {
        r#"<?php
#[Attribute]
class MetaAttr {
    public string $data;
    public function __construct(string $data) {
        $this->data = $data;
    }
}

#[MetaAttr('payload')]
class Annotated {}

$rc = new ReflectionClass(Annotated::class);
$attr = $rc->getAttributes()[0];
$instance = $attr->newInstance();

echo get_class($instance) . ":" . $instance->data;
"#,
        ["MetaAttr:payload"]
    };

    reflection_attribute_new_instance_named_args => {
        r#"<?php
#[Attribute]
class ConfigAttr {
    public function __construct(public bool $enabled, public int $level = 1) {}
}

#[ConfigAttr(level: 5, enabled: true)]
class Service {}

$rc = new ReflectionClass(Service::class);
$attr = $rc->getAttributes()[0];
$instance = $attr->newInstance();

echo $instance->enabled ? "yes" : "no";
echo "|" . $instance->level;
"#,
        ["yes|5"]
    };
}
