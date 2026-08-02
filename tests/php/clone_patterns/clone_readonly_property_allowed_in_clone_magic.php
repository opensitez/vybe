<?php
// vybe-test: php/clone_patterns/clone_readonly_property_allowed_in_clone_magic
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Token {
    public readonly string $value;
    public function __construct(string $v) { $this->value = $v; }
    public function withValue(string $v): static {
        $clone = clone $this;
        return $clone;
    }
}
$t = new Token("abc");
$t2 = $t->withValue("xyz");
echo $t->value;

__vybe_check(ob_get_clean(), "abc");
