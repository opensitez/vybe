<?php
// vybe-test: php/compact_extract/extract_prefix_conflicts_only
// origin: languages/php/tests/php/test_compact_extract.rs

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

$name = 'existing';
extract(['name' => 'new', 'age' => 25], EXTR_PREFIX_SAME, 'pre');
echo $name . ':' . $pre_name . ':' . $age;

__vybe_check(ob_get_clean(), "existing:new:25");
