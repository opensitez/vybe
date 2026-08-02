<?php
// vybe-test: php/callables/closure_bind_derived_object_base_scope
// origin: languages/php/tests/php/test_callables.rs

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

class Base { private int $v = 10; }
class Child extends Base {}
$read = Closure::bind(function(): int { return $this->v; }, new Child(), Base::class);
echo $read();

__vybe_check(ob_get_clean(), "10");
