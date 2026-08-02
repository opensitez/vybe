<?php
// vybe-test: php/string_functions_extended/preg_quote_preserves_custom_delimiter_escape_runtime
// origin: languages/php/tests/php/test_string_functions_extended.rs

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

$p = preg_quote('/a$b[c]', '/');
echo $p === "\/a\$b\[c\]" ? 'ok' : 'bad';
echo '|';
echo preg_match("/" . $p . "/", '/a$b[c]') === 1 ? 'match' : 'nomatch';

__vybe_check(ob_get_clean(), "ok|match");
