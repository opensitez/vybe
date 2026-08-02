<?php
// vybe-test: php/oop/trait_abstract_and_concrete_resolution_runtime
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

trait Renderer {
    abstract public function source(): string;
    public function wrap(): string {
        return '[' . $this->source() . ']';
    }
}
class Message {
    use Renderer;
    public function source(): string { return 'hi'; }
}
echo (new Message())->wrap();

__vybe_check(ob_get_clean(), "[hi]");
