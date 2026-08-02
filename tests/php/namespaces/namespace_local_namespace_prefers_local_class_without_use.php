<?php
// vybe-test: php/namespaces/namespace_local_namespace_prefers_local_class_without_use
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

namespace Shared {
    class Logger { public function channel(): string { return 'shared'; } }
}
namespace App {
    class Logger { public function channel(): string { return 'app'; } }
    $local = new Logger();
    $global = new \Shared\Logger();
    echo $local->channel() . '|' . $global->channel();
}

__vybe_check(ob_get_clean(), "app|shared");
