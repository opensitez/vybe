<?php
// vybe-test: php/method_chaining/chain_with_void_step_then_tail_method
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

class Pipeline {
    private string $log = '';
    public function touch(string $chunk): void {
        $this->log .= $chunk;
    }
    public function chain(string $chunk): static {
        $this->touch($chunk);
        return $this;
    }
    public function snapshot(): string { return $this->log; }
}
echo (new Pipeline())->chain('a')->chain('b')->snapshot();

__vybe_check(ob_get_clean(), "ab");
