<?php
// vybe-test: php/php_reflection_attribute_target_flags/test_reflection_attribute_target_class
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

#[Attribute(Attribute::TARGET_CLASS)]
class CustomAttr {}

#[CustomAttr]
class TargetClass {}

$rc = new ReflectionClass(TargetClass::class);
$attrs = $rc->getAttributes(CustomAttr::class);
echo count($attrs) . ':' . $attrs[0]->getName(), "\n";

__vybe_check(ob_get_clean(), "1:CustomAttr");
