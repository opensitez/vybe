<?php
// vybe-test: php/classes/class_magic_setter_getter_runtime
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

class Bag {
    private array $values = [];
    public function __set(string $name, mixed $value): void { $this->values[$name] = $value; }
    public function __get(string $name): mixed { return $this->values[$name] ?? null; }
}
$b = new Bag();
$b->lang = 'php';
echo $b->lang;

__vybe_check(ob_get_clean(), "php");
