<?php
// vybe-test: php/oop_interfaces/interface_const_aliasing_from_implementer_runtime
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Versioned { public const VERSION = '1'; }
class Plugin implements Versioned {
    public const VERSION = '2';
}
echo Plugin::VERSION;
echo '|';
echo Versioned::VERSION;

__vybe_check(ob_get_clean(), "2|1");
