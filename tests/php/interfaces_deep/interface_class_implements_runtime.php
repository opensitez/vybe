<?php
// vybe-test: php/interfaces_deep/interface_class_implements_runtime
// origin: languages/php/tests/php/test_interfaces_deep.rs

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

interface A {}
interface B {}
class C implements A, B {}
echo array_key_exists('A', class_implements(C::class)) ? 'A' : 'X';
echo '|';
echo array_key_exists('B', class_implements(C::class)) ? 'B' : 'X';

__vybe_check(ob_get_clean(), "A|B");
