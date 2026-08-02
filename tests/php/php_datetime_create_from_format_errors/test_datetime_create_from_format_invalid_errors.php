<?php
// vybe-test: php/php_datetime_create_from_format_errors/test_datetime_create_from_format_invalid_errors
// origin: languages/php/tests/php/test_php_datetime_create_from_format_errors.rs

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

$res = DateTime::createFromFormat('Y-m-d', 'invalid-date');
$errors = DateTime::getLastErrors();
echo ($res === false && is_array($errors) && ($errors['error_count'] > 0 || $errors['warning_count'] > 0 || count($errors['errors']) > 0)) ? 'errors_logged' : 'err', "\n";

__vybe_check(ob_get_clean(), "errors_logged");
