<?php
// vybe-test: php/advanced_oop/__call_static_dispatch_runtime
// origin: languages/php/tests/php/test_advanced_oop.rs

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

class HandlerRegistry {
    public static function __callStatic(string $name, array $args): mixed {
        return match($name) {
            'make' => 'created:' . ($args[0] ?? 'default'),
            default => null,
        };
    }
}
echo HandlerRegistry::make('report');

__vybe_check(ob_get_clean(), "created:report");
