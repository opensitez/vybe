<?php
// vybe-test: php/string_escaping/html_entity_decode_preserves_double_quotes_when_ent_noquotes
// origin: languages/php/tests/php/test_string_escaping.rs

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

$out = htmlspecialchars('a "b"', ENT_NOQUOTES);
echo str_contains($out, '&quot;') ? 'quoted' : 'no-quote';

__vybe_check(ob_get_clean(), "no-quote");
