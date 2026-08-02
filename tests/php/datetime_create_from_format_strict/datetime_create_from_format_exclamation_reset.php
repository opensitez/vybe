<?php
// vybe-test: php/datetime_create_from_format_strict/datetime_create_from_format_exclamation_reset
// origin: languages/php/tests/php/test_datetime_create_from_format_strict.rs

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

$base = new DateTime('1999-12-31 11:22:33');
$dt = DateTime::createFromFormat('!Y-m-d', '2000-01-02', $base->getTimezone());
echo $dt->format('Y-m-d H:i:s');

__vybe_check(ob_get_clean(), "2000-01-02 00:00:00");
