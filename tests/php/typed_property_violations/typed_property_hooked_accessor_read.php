<?php
// vybe-test: php/typed_property_violations/typed_property_hooked_accessor_read
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

class Meter {
    private int $v = 0;
    public int $reading {
        get => $this->v;
        set => $this->v = $value;
    }
}
$m = new Meter();
$m->reading = 4;
echo $m->reading;

__vybe_check(ob_get_clean(), "4");
