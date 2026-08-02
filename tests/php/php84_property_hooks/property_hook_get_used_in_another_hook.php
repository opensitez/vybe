<?php
// vybe-test: php/php84_property_hooks/property_hook_get_used_in_another_hook
// origin: languages/php/tests/php/test_php84_property_hooks.rs

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

class Measure {
    public function __construct(public float $meters) {}
    public float $cm { get => $this->meters * 100; }
    public float $mm { get => $this->cm * 10; }
}
$m = new Measure(1.5);
echo $m->cm . ',' . $m->mm;

__vybe_check(ob_get_clean(), "150,1500");
