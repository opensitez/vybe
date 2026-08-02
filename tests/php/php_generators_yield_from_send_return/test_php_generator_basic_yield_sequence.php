<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_basic_yield_sequence
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs

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

function rangeGen($start, $end) {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}

$items = [];
foreach (rangeGen(1, 4) as $num) {
    $items[] = $num;
}
echo implode("-", $items);

__vybe_check(ob_get_clean(), "1-2-3-4");
