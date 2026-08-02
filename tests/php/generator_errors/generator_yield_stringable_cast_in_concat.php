<?php
// vybe-test: php/generator_errors/generator_yield_stringable_cast_in_concat
// origin: languages/php/tests/php/test_generator_errors.rs

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

class Label { public function __construct(private string $t) {} public function __toString(): string { return $this->t; } }
function g(): Generator { yield new Label('z'); }
$g = g();
$g->next();
echo (string)$g->current();

__vybe_check(ob_get_clean(), "z");
