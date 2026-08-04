<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_relative_resolution_with_class_exists
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

namespace Framework\Core;
class Kernel {
    public static function name(): string { return 'kernel'; }
}

namespace App\Runtime;
echo class_exists('Kernel') ? 'inner' : 'miss';
echo '|';
echo class_exists('\\Framework\\Core\\Kernel') ? 'absolute' : 'noabs';
echo '|';
echo \Framework\Core\Kernel::name();

__vybe_check(ob_get_clean(), "miss|absolute|kernel");
