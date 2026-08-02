<?php
// vybe-test: php/datetime/test_date_parse_strict_and_error_state
// origin: languages/php/tests/php/test_datetime.rs

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

$d = date_parse('2024-12-01 12:00:00');
echo is_array($d) ? 'ok' : 'bad';
echo $d['error_count'];
echo '|';
$bad = date_parse('2024-13-99');
echo is_array($bad['errors']) ? 'errors' : 'noerr';

__vybe_check(ob_get_clean(), "ok0|errors");
