<?php
// vybe-test: php/traits/trait_static_method_callable
// origin: languages/php/tests/php/test_traits.rs

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

trait Factory {
    public static function make(): string { return 'new'; }
}
class Item { use Factory; }
echo Item::make();

__vybe_check(ob_get_clean(), "new");
