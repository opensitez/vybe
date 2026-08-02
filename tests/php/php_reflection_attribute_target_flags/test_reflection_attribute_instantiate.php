<?php
// vybe-test: php/php_reflection_attribute_target_flags/test_reflection_attribute_instantiate
// origin: languages/php/tests/php/test_php_reflection_attribute_target_flags.rs

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
class ParamAttr {
    public function __construct(public string $label) {}
}

#[ParamAttr('demo_label')]
class Demo {}

$rc = new ReflectionClass(Demo::class);
$attr = $rc->getAttributes(ParamAttr::class)[0]->newInstance();
echo $attr->label, "\n";

__vybe_check(ob_get_clean(), "demo_label");
