<?php
// vybe-test: php/class_instantiation/constructor_can_instantiate_class_declared_later
// origin: languages/php/tests/php/test_class_instantiation.rs

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

class Project {
    public function __construct() { $this->workflow = new Workflow(); }
    public function label(): string { return $this->workflow->label(); }
}
class Workflow {
    public function label(): string { return 'W'; }
}
$p = new Project();
echo $p->label();

__vybe_check(ob_get_clean(), "W");
