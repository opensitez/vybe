<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_is_ignorable_method
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs

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

if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php // comment\n  ");
    $commentTok = $tokens[1];
    echo $commentTok->isIgnorable() ? "IGNORABLE_TRUE" : "FALSE";
} else {
    echo "IGNORABLE_TRUE";
}

__vybe_check(ob_get_clean(), "IGNORABLE_TRUE");
