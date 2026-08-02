<?php
// vybe-test: php/generators_advanced/generator_pipeline_map_filter
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function rangeGen(int $start, int $end) {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
function filterGen($gen, callable $pred) {
    foreach ($gen as $val) {
        if ($pred($val)) yield $val;
    }
}
function mapGen($gen, callable $fn) {
    foreach ($gen as $val) {
        yield $fn($val);
    }
}
$numbers = rangeGen(1, 10);
$evens = filterGen($numbers, fn($n) => $n % 2 == 0);
$doubled = mapGen($evens, fn($n) => $n * 2);
$result = [];
foreach ($doubled as $v) {
    $result[] = $v;
}
echo implode(",", $result);

__vybe_check(ob_get_clean(), "4,8,12,16,20");
