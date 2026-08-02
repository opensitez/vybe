<?php
// vybe-test: php/namespaces/namespace_rootless_class_lookup_in_parent_segment
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

namespace App\Core {
    class Engine {
        public function version(): string { return 'v1'; }
    }
}
namespace App {
    function getEngineVersion(): string {
        $e = new Core\Engine();
        return $e->version();
    }
}
echo \App\getEngineVersion();

__vybe_check(ob_get_clean(), "v1");
