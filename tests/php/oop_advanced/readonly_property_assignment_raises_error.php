<?php
// vybe-test: php/oop_advanced/readonly_property_assignment_raises_error
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

class Snapshot {
    public readonly string $name;
    public function __construct(string $name) { $this->name = $name; }
}
$s = new Snapshot("init");
try {
    $s->name = "bad";
    echo "changed";
} catch (Error $e) {
    echo "readonly";
}

__vybe_check(ob_get_clean(), "readonly");
