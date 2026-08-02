<?php
// vybe-test: php/array_offset_access/count_on_non_countable_string_throws
// origin: languages/php/tests/php/test_array_offset_access.rs

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

try { count('abc'); echo 'ok'; }
catch (TypeError $e) { echo 'count-str'; }
catch (ValueError $e) { echo 'count-str'; }

__vybe_check(ob_get_clean(), "count-str");
