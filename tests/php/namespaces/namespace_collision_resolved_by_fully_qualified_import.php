<?php
// vybe-test: php/namespaces/namespace_collision_resolved_by_fully_qualified_import
// origin: languages/php/tests/php/test_namespaces.rs

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

namespace Legacy { class Logger { public function id(): string { return 'L'; } } }
namespace Modern { class Logger { public function id(): string { return 'M'; } } }
namespace App {
    use Legacy\Logger;
    function pick(): string {
        $old = new Logger();
        $new = new \Modern\Logger();
        return $old->id() . $new->id();
    }
}
echo \App\pick();

__vybe_check(ob_get_clean(), "LM");
