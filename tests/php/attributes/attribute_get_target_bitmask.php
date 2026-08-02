<?php
// vybe-test: php/attributes/attribute_get_target_bitmask
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

#[Attribute(Attribute::TARGET_CLASS | Attribute::TARGET_METHOD)]
class DualTarget {}
$rc = new ReflectionClass(DualTarget::class);
$attr = $rc->getAttributes(Attribute::class)[0]->newInstance();
echo ($attr->flags & Attribute::TARGET_CLASS) ? 'has_class_target' : 'err';

__vybe_check(ob_get_clean(), "has_class_target");
