<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_multiple_aliases_and_static_function_call
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

namespace Src\Lib;
class Fallback {
    public static function name(): string { return 'class'; }
}
function action(): string { return 'fn'; }

namespace App;
use Src\Lib\Fallback as AliasClass;
use function Src\Lib\action as alias_fn;
$kind = AliasClass::name();
echo $kind . '|' . alias_fn();

__vybe_check(ob_get_clean(), "class|fn");
