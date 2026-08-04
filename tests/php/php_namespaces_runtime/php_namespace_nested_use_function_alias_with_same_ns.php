<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_nested_use_function_alias_with_same_ns
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

namespace Payments\Stripe;
function status(): string { return 'stripe'; }

namespace Payments;
use Stripe\status as stripe_status;
echo stripe_status();

__vybe_check(ob_get_clean(), "stripe");
