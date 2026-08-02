<?php
// vybe-test: php/reflection_class_attributes/reflection_class_get_attributes_with_arguments
// origin: languages/php/tests/php/test_reflection_class_attributes.rs

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
class ArgAttr {
    public function __construct(public string $val) {}
}

#[ArgAttr('hello')]
class TargetClassArgs {}

$rc = new ReflectionClass(TargetClassArgs::class);
$attrs = $rc->getAttributes();
$args = $attrs[0]->getArguments();
echo $args[0];

__vybe_check(ob_get_clean(), "hello");
