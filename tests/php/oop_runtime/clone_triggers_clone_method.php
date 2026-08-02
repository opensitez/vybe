<?php
// vybe-test: php/oop_runtime/clone_triggers_clone_method
// origin: languages/php/tests/php/test_oop_runtime.rs

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

class Tag {
    public function __construct(public string $v) {}
    public function __clone(): void { $this->v .= '!'; }
}
$t = new Tag('x');
$c = clone $t;
echo $c->v;

__vybe_check(ob_get_clean(), "x!");
