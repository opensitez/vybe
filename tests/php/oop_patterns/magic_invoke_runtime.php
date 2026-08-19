<?php
// vybe-test: php/oop_patterns/magic_invoke_runtime
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

class CallableCounter {
    public function __construct(private int $n = 0) {}
    public function __invoke(int $step = 1): int {
        $this->n += $step;
        return $this->n;
    }
}
$c = new CallableCounter(2);
echo $c();
echo $c(3);

__vybe_check(ob_get_clean(), "36");
