<?php
// vybe-test: php/unserialize_errors/unserialize_object_sleep_controls_fields
// origin: languages/php/tests/php/test_unserialize_errors.rs

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

class Pick {
    public int $keep = 1;
    private int $hide = 9;
    public function __sleep(): array { return ['keep']; }
}
$o = unserialize(serialize(new Pick()));
echo $o->keep;

__vybe_check(ob_get_clean(), "1");
