<?php
// vybe-test: php/method_chaining/chain_with_nullsafe_operator_and_fallback
// origin: languages/php/tests/php/test_method_chaining.rs

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

class MaybeChain {
    public function step(): static {
        return $this;
    }
    public function value(): int {
        return 42;
    }
}

$obj = new MaybeChain();
$present = $obj?->step()?->value();
$absent = null?->step()?->value();
echo ($present === 42 ? 'yes' : 'no') . '|' . ($absent === null ? 'null' : 'val');

__vybe_check(ob_get_clean(), "yes|null");
