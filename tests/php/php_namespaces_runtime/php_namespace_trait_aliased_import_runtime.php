<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_trait_aliased_import_runtime
// origin: languages/php/tests/php/test_php_namespaces_runtime.rs

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

namespace Traits;
trait Logger {
    public function log(): string { return 'ok'; }
}

namespace App;
use Traits\Logger as AppLogger;
class Service {
    use AppLogger { log as public emit; }
}
echo (new Service())->emit();

__vybe_check(ob_get_clean(), "ok");
