<?php
// vybe-test: php/generators_send_throw/infinite_generator_take_n_items
// origin: languages/php/tests/php/test_generators_send_throw.rs

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

function naturals(): Generator {
    $n = 1;
    while (true) { yield $n++; }
}
$result = [];
foreach (naturals() as $v) {
    $result[] = $v;
    if (count($result) >= 5) break;
}
echo implode(',', $result);

__vybe_check(ob_get_clean(), "1,2,3,4,5");
