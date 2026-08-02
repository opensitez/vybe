<?php
// vybe-test: php/method_chaining/chain_with_magic_call_proxy
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

class Proxy {
    private int $value = 0;
    public function __call(string $name, array $args): static {
        if ($name === 'add') { $this->value += (int)($args[0] ?? 0); }
        return $this;
    }
    public function value(): int { return $this->value; }
}
echo (new Proxy())->add(2)->add(3)->value();

__vybe_check(ob_get_clean(), "5");
