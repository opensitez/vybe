<?php
// vybe-test: php/scope_global_static/static_property_isolation_in_methods
// origin: languages/php/tests/php/test_scope_global_static.rs

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

class Counter {
    public function inc(): int { static $n = 0; return ++$n; }
}
$a = new Counter();
$b = new Counter();
echo $a->inc() . $a->inc();
echo '|';
echo $b->inc();

__vybe_check(ob_get_clean(), "12|3");
