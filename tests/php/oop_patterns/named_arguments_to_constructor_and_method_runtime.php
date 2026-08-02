<?php
// vybe-test: php/oop_patterns/named_arguments_to_constructor_and_method_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class Metrics {
    public function __construct(public int $a = 0, public int $b = 0) {}
    public function span(int $start = 0, int $end = 0): int {
        return $end - $start;
    }
}
$m = new Metrics(b: 30, a: 10);
echo $m->span(end: 15, start: 5) . '|' . $m->a . ':' . $m->b;

__vybe_check(ob_get_clean(), "10|10:30");
