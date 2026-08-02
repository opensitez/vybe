<?php
// vybe-test: php/method_chaining/chain_with_method_and_array_key_access
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

class Record {
    private array $items = [];
    public function set(string $k, int $v): static { $this->items[$k] = $v; return $this; }
    public function getAll(): array { return $this->items; }
}
$record = (new Record())->set('first', 1)->set('second', 2);
echo $record->getAll()['second'];

__vybe_check(ob_get_clean(), "2");
