<?php
// vybe-test: php/php_object_chaining/php_magic_chain_runtime
// origin: languages/php/tests/php/test_php_object_chaining.rs

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

class Counter {
    public function __construct(private int $value) {}
    public function inc(): self { $this->value += 1; return $this; }
    public function __get(string $name): mixed { return $this->value; }
}
$c = (new Counter(0))->inc()->inc();
echo $c->value;

__vybe_check(ob_get_clean(), "2");
