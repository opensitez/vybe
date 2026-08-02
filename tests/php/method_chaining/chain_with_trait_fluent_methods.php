<?php
// vybe-test: php/method_chaining/chain_with_trait_fluent_methods
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

trait TChain {
    public function add(int $v): static {
        if (!isset($this->items)) {
            $this->items = [];
        }
        $this->items[] = $v;
        return $this;
    }
}
class Basket {
    use TChain;
    public array $items = [];
    public function total(): int { return array_sum($this->items); }
}
echo (new Basket())->add(2)->add(3)->total();

__vybe_check(ob_get_clean(), "5");
