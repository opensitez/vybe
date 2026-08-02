<?php
// vybe-test: php/php_class_alias_autoload_parameter/test_class_alias_autoload_true
// origin: languages/php/tests/php/test_php_class_alias_autoload_parameter.rs

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

class OriginalClass {
    public function identity(): string { return "original"; }
}
class_alias('OriginalClass', 'AliasedClass', true);
$obj = new AliasedClass();
echo $obj->identity(), "\n";

__vybe_check(ob_get_clean(), "original");
