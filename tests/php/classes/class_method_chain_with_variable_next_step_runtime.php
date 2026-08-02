<?php
// vybe-test: php/classes/class_method_chain_with_variable_next_step_runtime
// origin: languages/php/tests/php/test_classes.rs

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

class Pipeline {
    private int $v = 0;
    public function step1(int $n): self { $this->v += $n; return $this; }
    public function step2(string $label): string { return $label . ':' . $this->v; }
}
$p = new Pipeline();
$next = 'step1';
echo $p->$next(3)->step2('ok');

__vybe_check(ob_get_clean(), "ok:3");
