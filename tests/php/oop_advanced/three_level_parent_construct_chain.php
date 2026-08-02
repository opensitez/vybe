<?php
// vybe-test: php/oop_advanced/three_level_parent_construct_chain
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class A {
    protected string $log = "";
    public function __construct() {
        $this->log .= "A";
    }
}
class B extends A {
    public function __construct() {
        parent::__construct();
        $this->log .= "B";
    }
}
class C extends B {
    public function __construct() {
        parent::__construct();
        $this->log .= "C";
    }
    public function getLog(): string { return $this->log; }
}
$c = new C();
echo $c->getLog(), "\n";

__vybe_check(ob_get_clean(), "ABC");
