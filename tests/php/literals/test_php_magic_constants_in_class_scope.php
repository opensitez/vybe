<?php
// vybe-test: php/literals/test_php_magic_constants_in_class_scope
// origin: languages/php/tests/php/test_literals.rs

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

namespace NS;
class Sample {
    public function identity(): string {
        return __CLASS__ . ':' . __FUNCTION__ . ':' . __METHOD__ . ':' . __NAMESPACE__ . ':' . __TRAIT__;
    }
}
$obj = new Sample();
echo $obj->identity();

__vybe_check(ob_get_clean(), "NS\\Sample:identity:Sample::identity:NS:");
