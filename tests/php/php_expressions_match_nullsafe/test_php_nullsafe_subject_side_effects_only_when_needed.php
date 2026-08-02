<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_nullsafe_subject_side_effects_only_when_needed
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs

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

class Node {
    public ?Node $next = null;
    public function label(): string { return 'leaf'; }
}
$root = null;
$count = 0;
$value = $root?->next?->label();
echo match ($value ?? 'fallback') {
    'leaf' => 'got',
    'fallback' => 'miss',
    default => 'other',
};
echo '|';
$root = new Node();
$root->next = new Node();
$value2 = $root?->next?->label();
echo match ($value2 ?? 'fallback') {
    'leaf' => 'got2',
    'fallback' => 'miss2',
    default => 'other2',
};

__vybe_check(ob_get_clean(), "miss|got2");
