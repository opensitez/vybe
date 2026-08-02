<?php
// vybe-test: php/namespaces/namespace_function_alias_chain_with_local_shadow
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

namespace Utils {
    function ping(string $value): string { return "u:$value"; }
}
namespace App {
    function ping(string $value): string { return "a:$value"; }
    use function Utils\ping as util_ping;
    function run(string $value): string {
        return ping($value) . '|' . util_ping($value);
    }
    echo run('x');
}

__vybe_check(ob_get_clean(), "a:x|u:x");
