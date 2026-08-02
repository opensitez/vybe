<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_use_global_access_inside_nested_scope
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs

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

$g = 'outside';
$fn = function() {
    global $g;
    return $g;
};
$g = 'inside';
echo $fn();

__vybe_check(ob_get_clean(), "inside");
