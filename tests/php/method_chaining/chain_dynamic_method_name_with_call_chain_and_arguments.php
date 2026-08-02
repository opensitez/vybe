<?php
// vybe-test: php/method_chaining/chain_dynamic_method_name_with_call_chain_and_arguments
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

class Writer {
    private string $s = '';
    public function append(string $chunk): static { $this->s .= $chunk; return $this; }
    public function output(): string { return $this->s; }
}
$writer = new Writer();
$method = 'append';
echo $writer->{$method}('a')->{$method}('b')->output();

__vybe_check(ob_get_clean(), "ab");
