<?php
// vybe-test: php/mb_substitute_character_fallback/mb_substitute_character_set_get
// origin: languages/php/tests/php/test_mb_substitute_character_fallback.rs

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

$old = mb_substitute_character();
mb_substitute_character(0x3013);
echo mb_substitute_character() . "|";
mb_substitute_character("none");
echo mb_substitute_character() . "|";
mb_substitute_character($old);
echo "restored";

__vybe_check(ob_get_clean(), "12307|none|restored");
