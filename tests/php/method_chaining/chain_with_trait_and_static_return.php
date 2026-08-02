<?php
// vybe-test: php/method_chaining/chain_with_trait_and_static_return
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

trait FluentMath {
    public function addOne(): static {
        $this->value += 1;
        return $this;
    }
}
class Counter {
    use FluentMath;
    public int $value = 0;
    public function add(int $v): static {
        $this->value += $v;
        return $this;
    }
}
echo (new Counter())->add(2)->addOne()->add(3)->addOne()->value;

__vybe_check(ob_get_clean(), "7");
