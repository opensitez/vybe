<?php
// vybe-test: php/method_chaining/chain_bool_flags_then_render
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

class Flags {
    private int $mask = 0;
    public function on(int $bit): static { $this->mask |= $bit; return $this; }
    public function has(int $bit): bool { return ($this->mask & $bit) !== 0; }
}
$f = (new Flags())->on(1)->on(4);
echo ($f->has(1) ? 'a' : '') . ($f->has(4) ? 'b' : '');

__vybe_check(ob_get_clean(), "ab");
