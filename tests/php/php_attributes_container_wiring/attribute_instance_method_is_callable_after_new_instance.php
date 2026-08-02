<?php
// vybe-test: php/php_attributes_container_wiring/attribute_instance_method_is_callable_after_new_instance
// origin: languages/php/tests/php/test_php_attributes_container_wiring.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

#[Attribute]
class Transform {
    public function __construct(private string $prefix) {}
    public function apply(string $v): string { return $this->prefix . $v; }
}
class Field {
    #[Transform('pre-')]
    public $name;
}
$t = (new ReflectionProperty(Field::class, 'name'))->getAttributes(Transform::class)[0]->newInstance();
echo $t->apply('value');

__vybe_check(ob_get_clean(), "pre-value");
