<?php
// vybe-test: php/oop/static_property_in_trait_runtime
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

trait Counters {
    public static int $n = 0;
    public static function bump(): int { return ++self::$n; }
}
class Widget {
    use Counters;
}
echo Widget::bump();
echo Widget::bump();
echo Widget::$n;

__vybe_check(ob_get_clean(), "122");
