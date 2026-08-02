<?php
// vybe-test: php/attributes/attribute_on_promoted_constructor_parameter
// origin: languages/php/tests/php/test_attributes.rs

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
class Inject {}
class Service {
    public function __construct(#[Inject] public string $name) {}
}
$ref = (new ReflectionClass(Service::class))->getConstructor()->getParameters()[0];
echo $ref->getAttributes(Inject::class) ? 'inject' : 'plain';

__vybe_check(ob_get_clean(), "inject");
