<?php
// vybe-test: php/error_handler/handler_records_errno_as_hex_suffix
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

$out = [];
set_error_handler(function(int $no, string $msg) use (&$out): bool {
    $out[] = dechex($no) . ':' . $msg;
    return true;
});
trigger_error('warn', E_USER_WARNING);
restore_error_handler();
echo implode('|', $out);

__vybe_check(ob_get_clean(), "200:warn");
