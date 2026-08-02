<?php
// vybe-test: php/datetime/datetime_parse_and_checkruntime
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

$p = date_parse('2024-11-30 10:20:30');
echo is_array($p) ? 'ok' : 'bad';
echo '|';
echo isset($p['warning_count']) ? 'warn' . $p['warning_count'] : 'nowarn';

__vybe_check(ob_get_clean(), "ok|warn0");
