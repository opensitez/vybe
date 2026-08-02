<?php
// vybe-test: php/generators_advanced/generator_with_default_parameter
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

function counter(int $start = 0, int $step = 1) {
    $n = $start;
    while (true) {
        yield $n;
        $n += $step;
    }
}
$g = counter(10, 5);
$result = [];
for ($i = 0; $i < 4; $i++) {
    $result[] = $g->current();
    $g->next();
}
echo implode(",", $result);

__vybe_check(ob_get_clean(), "10,15,20,25");
