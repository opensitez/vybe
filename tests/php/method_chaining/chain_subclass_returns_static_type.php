<?php
// vybe-test: php/method_chaining/chain_subclass_returns_static_type
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

class Base {
    protected string $tag = '';
    public function mark(string $t): static { $this->tag = $t; return $this; }
    public function read(): string { return $this->tag; }
}
class Derived extends Base {}
echo (new Derived())->mark('d')->read();

__vybe_check(ob_get_clean(), "d");
