<?php
// vybe-test: php/method_chaining/chain_terminates_with_array_access_expression
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

class Rows {
    private array $rows = [];
    public function add(array $row): static { $this->rows[] = $row; return $this; }
    public function rows(): array { return $this->rows; }
}
echo (new Rows())->add(['id' => 1])->add(['id' => 2])->rows()[1]['id'];

__vybe_check(ob_get_clean(), "2");
