<?php
// vybe-test: php/reflection_attribute_instantiation/reflection_attribute_new_instance
// origin: languages/php/tests/php/test_reflection_attribute_instantiation.rs

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

__vybe_check(ob_get_clean(), "MetaAttr:payload");
