<?php
// vybe-test: php/method_chaining/chain_with_conditional_method_chain
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

class Guarded {
    private int $n = 0;
    public function add(int $x): static { $this->n += $x; return $this; }
    public function maybe(bool $ok): ?self { return $ok ? $this : null; }
    public function value(): int { return $this->n; }
}
echo ($ok = (new Guarded())->add(1)->maybe(true)->add(2)->value());
echo '|' . ((new Guarded())->maybe(false)?->add(3)?->value() ?? 99);

__vybe_check(ob_get_clean(), "3|99");
