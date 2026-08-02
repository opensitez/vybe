<?php
// vybe-test: php/compact_extract/extract_prefix_all
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

extract(['id' => 1, 'name' => 'Joe'], EXTR_PREFIX_ALL, 'usr');
echo $usr_id . ':' . $usr_name;

__vybe_check(ob_get_clean(), "1:Joe");
