<?php
// vybe-test: php/php_script_metadata_getlastmod/test_getlastmod_returns_timestamp_or_false
// origin: languages/php/tests/php/test_php_script_metadata_getlastmod.rs

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

$mod = getlastmod();
echo ($mod === false || (is_int($mod) && $mod > 0)) ? 'lastmod_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "lastmod_ok");
