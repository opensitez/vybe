<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_forward_static_calls_static_factory_chain
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs

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

class ParentFactory {
    public static function make(): static {
        return new static();
    }

    public static function label(): string {
        return "parent";
    }
}

class ChildFactory extends ParentFactory {
    public static function label(): string {
        return "child";
    }
}

$factory = ChildFactory::make();
echo get_class($factory) . "|" . $factory::label();

__vybe_check(ob_get_clean(), "ChildFactory|child");
