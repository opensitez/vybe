<?php
// vybe-test: php/method_chaining/chain_with_property_accessor_chain
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

class Payload {
    public function __construct(private array $state = []) {}
    public function put(string $k, string $v): static {
        $this->state[$k] = $v;
        return $this;
    }
    public function state(): array { return $this->state; }
}
echo (new Payload())->put('a', 'x')->put('b', 'y')->state()['b'];

__vybe_check(ob_get_clean(), "y");
