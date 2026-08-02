<?php
// vybe-test: php/array_uintersect_uassoc_callback/array_uintersect_uassoc_value_callback_exception
// origin: languages/php/tests/php/test_array_uintersect_uassoc_callback.rs

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

$a = ["a" => "A", "b" => "B"];
$b = ["a" => "a"];
try {
    array_uintersect_uassoc(
        $a,
        $b,
        function($v1, $v2) {
            throw new RuntimeException('intersect-failed');
        },
        fn($k1, $k2) => strcmp((string)$k1, (string)$k2)
    );
    echo 'no-exception';
} catch (Throwable $e) {
    echo $e->getMessage();
}

__vybe_check(ob_get_clean(), "intersect-failed");
