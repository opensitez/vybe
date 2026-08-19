<?php
// vybe-test: php/method_chaining/chain_with_intermediate_object_reassignment
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

class Stage {
    private int $v = 0;
    public function add(int $x): static { $this->v += $x; return $this; }
    public function fork(): static {
        $next = new Stage();
        $next->add($this->v);
        return $next;
    }
    public function value(): int { return $this->v; }
}
$start = new Stage();
$final = $start->add(3)->fork()->add(4);
echo $start->value() . '|' . $final->value();

__vybe_check(ob_get_clean(), "3|7");
