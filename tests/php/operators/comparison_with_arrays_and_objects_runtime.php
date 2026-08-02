<?php
// vybe-test: php/operators/comparison_with_arrays_and_objects_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo ([] == []) ? 'array-eq' : 'array-ne';
echo '|';
echo ([] === []) ? 'array-ident' : 'array-not-ident';
echo '|';
class O {}
$a = new O();
$b = new O();
echo ($a == $b) ? 'obj-eq' : 'obj-ne';
echo '|';
echo ($a === $b) ? 'obj-id' : 'obj-not-id';

__vybe_check(ob_get_clean(), "array-eq|array-ident|obj-eq|obj-not-id");
