<?php
// vybe-test: php/type_hints_advanced/return_static_new_instance
// origin: languages/php/tests/php/test_type_hints_advanced.rs

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

class Fluent {
    private array $items = [];
    public function push(mixed $v): static { $this->items[] = $v; return $this; }
    public function count(): int { return count($this->items); }
}
$f = (new Fluent)->push(1)->push(2)->push(3);
echo $f->count();

__vybe_check(ob_get_clean(), "3");
