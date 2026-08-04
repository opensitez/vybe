<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_group_use_style_runtime
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

namespace Grouped\Storage;
class FileStore { public static function kind(): string { return 'file'; } }
function provider(): string { return 'provider'; }
const STORE = 'disk';

namespace App;
use Grouped\Storage\{FileStore, provider, STORE};
echo FileStore::kind() . '|' . provider() . '|' . STORE;

__vybe_check(ob_get_clean(), "file|provider|disk");
