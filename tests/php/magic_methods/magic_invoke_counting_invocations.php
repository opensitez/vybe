<?php
// vybe-test: php/magic_methods/magic_invoke_counting_invocations
// origin: languages/php/tests/php/test_magic_methods.rs

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

class CallCounter {
    private int $calls = 0;
    public function __invoke(int $x): int {
        $this->calls++;
        return $x * $this->calls;
    }
    public function getCalls(): int { return $this->calls; }
}
$fn = new CallCounter();
echo $fn(10);
echo $fn(10);
echo $fn(10);
echo $fn->getCalls();

__vybe_check(ob_get_clean(), "1020303");
