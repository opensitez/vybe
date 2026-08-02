<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_final_method_in_abstract_base_class
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs

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

abstract class TemplateMethodProcessor {
    final public function execute(): string {
        return "PRE -> " . $this->step() . " -> POST";
    }
    abstract protected function step(): string;
}

class ConcreteProcessor extends TemplateMethodProcessor {
    protected function step(): string { return "STEP_BODY"; }
}

$cp = new ConcreteProcessor();
echo $cp->execute();

__vybe_check(ob_get_clean(), "PRE -> STEP_BODY -> POST");
