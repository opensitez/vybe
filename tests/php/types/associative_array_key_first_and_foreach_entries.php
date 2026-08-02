<?php
// vybe-test: php/types/associative_array_key_first_and_foreach_entries
// origin: languages/php/tests/php/test_types.rs

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

$files = []; $files['a.knt'] = '/tmp/a'; $files['b.knt'] = '/tmp/b'; echo array_key_first($files); foreach ($files as $key => $path) { echo $key; echo $path; }

__vybe_check(ob_get_clean(), "a.knta.knt/tmp/ab.knt/tmp/b");
