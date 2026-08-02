<?php
// vybe-test: php/operators/null_coalesce_nested_array_indexing_chain_runtime
// origin: languages/php/tests/php/test_operators.rs

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

$payload = ['user' => ['profile' => null]];
echo $payload['user']['name'] ?? 'anon';
echo '|';
echo ($payload['user']['profile'] ?? $payload['user']['fallback'] ?? 'missing');
echo '|';
echo ($payload['team']['owner'] ?? 'team-unknown');

__vybe_check(ob_get_clean(), "anon|missing|team-unknown");
