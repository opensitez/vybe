<?php
// vybe-test: php/mb_convert_variables_references/mb_convert_variables_basic
// origin: languages/php/tests/php/test_mb_convert_variables_references.rs

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

$var1 = "äöü";
$var2 = ["äöü", "test"];
$enc = mb_convert_variables("UTF-8", "ISO-8859-1", $var1, $var2);
echo is_string($enc) ? "ok" : "fail";

__vybe_check(ob_get_clean(), "ok");
