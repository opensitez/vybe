<?php
// vybe-test: php/classes/class_abstract_trait_requirement_runtime
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

trait HasFormatter {
    abstract public function format(): string;
    public function output(): string { return '[' . $this->format() . ']'; }
}
class Formatted {
    use HasFormatter;
    public function __construct(public string $name) {}
    public function format(): string { return $this->name; }
}
echo (new Formatted('ok'))->output();

__vybe_check(ob_get_clean(), "[ok]");
