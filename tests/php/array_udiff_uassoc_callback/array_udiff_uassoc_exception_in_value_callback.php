<?php
// vybe-test: php/array_udiff_uassoc_callback/array_udiff_uassoc_exception_in_value_callback
// origin: languages/php/tests/php/test_array_udiff_uassoc_callback.rs

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

$a = ["a" => 1, "b" => 2];
$b = ["c" => 3];
try {
    array_udiff_uassoc(
        $a,
        $b,
        function($v1, $v2) {
            throw new RuntimeException('diff-failed');
        },
        fn($k1, $k2) => strcmp($k1, $k2)
    );
    echo 'no-exception';
} catch (Throwable $e) {
    echo $e->getMessage();
}

__vybe_check(ob_get_clean(), "no-exception");
