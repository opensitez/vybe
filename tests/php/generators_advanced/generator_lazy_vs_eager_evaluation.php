<?php
// vybe-test: php/generators_advanced/generator_lazy_vs_eager_evaluation
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

$calls_gen = 0;
$calls_arr = 0;
function lazyDoubles(array $items, int &$counter) {
    foreach ($items as $v) {
        $counter++;
        yield $v * 2;
    }
}
function eagerDoubles(array $items, int &$counter): array {
    $counter += count($items);
    return array_map(fn($v) => $v * 2, $items);
}
$data = [1, 2, 3, 4, 5];
$gen = lazyDoubles($data, $calls_gen);
// only consume first 2
$taken = [];
for ($i = 0; $i < 2; $i++) {
    $taken[] = $gen->current();
    $gen->next();
}
echo $calls_gen;     // only 2 calls made
$arr = eagerDoubles($data, $calls_arr);
echo $calls_arr;     // all 5 evaluated upfront
echo implode(",", $taken);

__vybe_check(ob_get_clean(), "352,4");
