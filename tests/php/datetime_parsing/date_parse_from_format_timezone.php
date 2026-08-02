<?php
// vybe-test: php/datetime_parsing/date_parse_from_format_timezone
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

$p = date_parse_from_format('d/m/Y H:i O', '15/08/2024 09:30 +0300');
echo $p['error_count'] === 0 ? 'ok' : 'bad';
echo '|';
echo isset($p['zone']) && is_string($p['zone']) ? 'zone' : 'nozone';

__vybe_check(ob_get_clean(), "ok|zone");
