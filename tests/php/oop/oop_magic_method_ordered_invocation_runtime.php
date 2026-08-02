<?php
// vybe-test: php/oop/oop_magic_method_ordered_invocation_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Recorder {
    private array $trace = [];
    public function __set(string $name, mixed $value): void {
        $this->trace[] = 'set:' . $name;
    }
    public function trace(): string { return implode(',', $this->trace); }
}
$r = new Recorder();
$r->v = 7;
echo $r->trace();

__vybe_check(ob_get_clean(), "set:v");
