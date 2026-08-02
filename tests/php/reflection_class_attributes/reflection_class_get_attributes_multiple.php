<?php
// vybe-test: php/reflection_class_attributes/reflection_class_get_attributes_multiple
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

__vybe_check(ob_get_clean(), "2|MyAttr1|MyAttr2");
