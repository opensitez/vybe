<?php
// vybe-test: php/oop_patterns/late_static_factory_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

abstract class BaseProduct {
    public function __construct(public string $name) {}
    public static function make(string $name): static {
        return new static($name);
    }
}

class Widget extends BaseProduct {}

$w = Widget::make('widget');
echo get_class($w) . '|' . $w->name;

__vybe_check(ob_get_clean(), "Widget|widget");
