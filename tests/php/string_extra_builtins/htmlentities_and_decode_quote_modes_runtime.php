<?php
// vybe-test: php/string_extra_builtins/htmlentities_and_decode_quote_modes_runtime
// origin: languages/php/tests/php/test_string_extra_builtins.rs

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

$encoded = htmlentities('<a>"&\'', ENT_QUOTES);
echo str_contains($encoded, '&lt;') ? 'lt' : 'no';
echo '|';
echo str_contains($encoded, '&quot;') ? 'dq' : 'no';
echo '|';
echo str_contains($encoded, '&#039;') ? 'sq' : 'no';
echo '|';
echo html_entity_decode($encoded, ENT_QUOTES);

__vybe_check(ob_get_clean(), "lt|dq|sq|<a>\"'");
