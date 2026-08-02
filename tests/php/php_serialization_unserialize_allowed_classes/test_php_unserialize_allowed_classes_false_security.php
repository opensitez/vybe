<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_unserialize_allowed_classes_false_security
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs

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

class Dangerous {
    public string $cmd = "rm -rf /";
}

$payload = serialize(new Dangerous());
$restored = unserialize($payload, ["allowed_classes" => false]);
echo is_object($restored) ? get_class($restored) : "NOT_OBJECT";

__vybe_check(ob_get_clean(), "__PHP_Incomplete_Class");
