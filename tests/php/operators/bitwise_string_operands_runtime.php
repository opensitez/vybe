<?php
// vybe-test: php/operators/bitwise_string_operands_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo ord(('a' | 'b')[0]) . '|';
echo ord(('foo' & 'bar')[0]) . ord(('foo' & 'bar')[1]) . ord(('foo' & 'bar')[2]) . '|';
echo ord(('abc' ^ 'bcd')[0]) . ord(('abc' ^ 'bcd')[1]) . ord(('abc' ^ 'bcd')[2]);

__vybe_check(ob_get_clean(), "99|989798|317");
