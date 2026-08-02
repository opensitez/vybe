<?php
// vybe-test: php/modern_php_deep/named_args_mixed_with_positional
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function slice(array $arr, int $offset, ?int $length = null, bool $preserve = false): array {
    return array_slice($arr, $offset, $length, $preserve);
}
$a = [1, 2, 3, 4, 5];
$r = slice($a, 1, length: 3);
echo implode(",", $r);

__vybe_check(ob_get_clean(), "2,3,4");
