<?php
// vybe-test: php/magic_methods/magic_call_passes_name_and_args
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Logger {
    private array $log = [];
    public function __call($method, $args) {
        $this->log[] = "$method(" . implode(",", $args) . ")";
    }
    public function dump(): string { return implode("|", $this->log); }
}
$l = new Logger();
$l->info("a");
$l->warn("b", "c");
echo $l->dump();

__vybe_check(ob_get_clean(), "info(a)|warn(b,c)");
