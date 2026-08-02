<?php
// vybe-test: php/oop/abstract_template_method_runtime
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

abstract class Template {
    abstract protected function body(): string;
    public function render(): string {
        return "<" . $this->body() . ">";
    }
}
class MessageTemplate extends Template {
    protected function body(): string { return "ok"; }
}
echo (new MessageTemplate())->render();

__vybe_check(ob_get_clean(), "<ok>");
