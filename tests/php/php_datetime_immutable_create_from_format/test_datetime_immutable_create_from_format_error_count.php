<?php
// vybe-test: php/php_datetime_immutable_create_from_format/test_datetime_immutable_create_from_format_error_count
// origin: languages/php/tests/php/test_php_datetime_immutable_create_from_format.rs

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

$dt = DateTimeImmutable::createFromFormat('Y-m-d', 'not-a-date');
echo $dt === false ? 'false' : 'ok';
$e = DateTimeImmutable::getLastErrors();
echo '|';
echo (($e['error_count'] ?? 0) > 0) ? 'errs' : 'clean';

__vybe_check(ob_get_clean(), "false|errs");
