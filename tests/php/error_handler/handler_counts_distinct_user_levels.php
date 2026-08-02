<?php
// vybe-test: php/error_handler/handler_counts_distinct_user_levels
// origin: languages/php/tests/php/test_error_handler.rs

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

$counts = ['w' => 0, 'n' => 0, 'd' => 0];
set_error_handler(function(int $no) use (&$counts): bool {
    if ($no === E_USER_WARNING) $counts['w']++;
    if ($no === E_USER_NOTICE) $counts['n']++;
    if ($no === E_USER_DEPRECATED) $counts['d']++;
    return true;
});
trigger_error('a', E_USER_WARNING);
trigger_error('b', E_USER_NOTICE);
trigger_error('c', E_USER_DEPRECATED);
trigger_error('d', E_USER_WARNING);
restore_error_handler();
echo $counts['w'] . $counts['n'] . $counts['d'];

__vybe_check(ob_get_clean(), "211");
