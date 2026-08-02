<?php
// vybe-test: php/datetime_parsing/datetime_modify_with_invalid_modifier_string_returns_false
// origin: languages/php/tests/php/test_datetime_parsing.rs

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

$d = new DateTime('2024-01-01');
$ok = $d->modify('not a valid modifier');
echo $ok === false ? 'bad-mod' : 'ok';

__vybe_check(ob_get_clean(), "bad-mod");
