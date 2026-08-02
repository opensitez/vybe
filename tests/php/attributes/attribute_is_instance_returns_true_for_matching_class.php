<?php
// vybe-test: php/attributes/attribute_is_instance_returns_true_for_matching_class
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
class Flag {}
#[Flag]
class Marked {}
$attr = (new ReflectionClass(Marked::class))->getAttributes()[0];
echo $attr->isInstance(Flag::class) ? 'flag' : 'other';

__vybe_check(ob_get_clean(), "flag");
