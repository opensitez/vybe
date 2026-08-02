<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_match_with_nested_nullsafe_subject
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs

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

class Node { public ?Node $next = null; public string $name = 'root'; }
$root = new Node();
$root->next = new Node();
$root->next->name = 'leaf';
$subject = $root->next?->name;
echo match ($subject) {
    null => 'missing',
    'leaf' => 'leaf',
    default => 'other',
};
echo '|';
echo match ($root->next?->next?->name) {
    null => 'no-child',
    default => 'has-child',
};

__vybe_check(ob_get_clean(), "leaf|no-child");
