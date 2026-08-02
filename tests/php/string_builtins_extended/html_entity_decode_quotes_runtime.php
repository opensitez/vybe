<?php
// vybe-test: php/string_builtins_extended/html_entity_decode_quotes_runtime
// origin: languages/php/tests/php/test_string_builtins_extended.rs

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

echo html_entity_decode("&lt;b&gt;A&amp;B&gt;&gt;", ENT_QUOTES);
echo "|";
echo html_entity_decode("&lt;i&gt;X&lt;/i&gt;");

__vybe_check(ob_get_clean(), "<b>A&B>>|<i>X</i>");
