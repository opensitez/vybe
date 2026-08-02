<?php
// vybe-test: php/namespaces/namespace_alias_collision_prefers_imported_alias
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

namespace Local {
    class Logger { public function tag(): string { return 'local'; } }
}
namespace Shared {
    class Logger { public function tag(): string { return 'shared'; } }
}
namespace App {
    use Shared\Logger;
    function pick(): string {
        $a = new Logger();
        $b = new \Local\Logger();
        return $a->tag() . '|' . $b->tag();
    }
}
echo \App\pick();

__vybe_check(ob_get_clean(), "shared|local");
