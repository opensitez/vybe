<?php
// vybe-test: php/static_closures/callable_type_hint_accepts_closure
// origin: languages/php/tests/php/test_static_closures.rs

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

function transform(callable $fn, array $items): array {
    return array_map($fn, $items);
}
$result = transform(static fn($x) => $x ** 2, [1, 2, 3, 4]);
echo implode(',', $result);

__vybe_check(ob_get_clean(), "1,4,9,16");
