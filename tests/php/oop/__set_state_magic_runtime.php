<?php
// vybe-test: php/oop/__set_state_magic_runtime
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

class State {
    public function __construct(public int $n) {}
    public static function __set_state(array $props): State {
        return new State($props['n'] + 1);
    }
}
$state = var_export(new State(3), true);
$obj = eval('return ' . $state . ';');
echo $obj->n;

__vybe_check(ob_get_clean(), "4");
