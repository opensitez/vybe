<?php
// vybe-test: php/oop/oop_clone_with_magic_reset_runtime
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

class Buffer {
    public string $value;
    public function __construct(string $value) { $this->value = $value; }
    public function __clone(): void { $this->value .= '-cloned'; }
}
$original = new Buffer('base');
$copy = clone $original;
echo $original->value;
echo '|';
echo $copy->value;

__vybe_check(ob_get_clean(), "base|base-cloned");
