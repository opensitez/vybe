<?php
// vybe-test: php/array_walk_recursive_by_ref/array_walk_recursive_can_set_null
// origin: languages/php/tests/php/test_array_walk_recursive_by_ref.rs

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

$a = ['x' => ['v' => 10], 'y' => ['v' => 20], 'z' => 'keep'];
array_walk_recursive($a, function(&$v) {
    if (is_int($v)) {
        $v = null;
    }
});
echo (($a['x']['v'] === null ? '1' : '0') . ($a['y']['v'] === null ? '1' : '0') . ($a['z'] === null ? '1' : '0'));

__vybe_check(ob_get_clean(), "110");
