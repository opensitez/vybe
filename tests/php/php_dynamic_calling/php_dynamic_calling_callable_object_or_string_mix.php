<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_callable_object_or_string_mix
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

$suffix = fn(string $s): string => $s . '!';
echo is_callable($suffix) ? 'func' : 'no';
echo '|';
echo $suffix('ok');
echo '|';
echo is_callable('strlen') ? 'strlen' : 'non';

__vybe_check(ob_get_clean(), "func|ok!|strlen");
