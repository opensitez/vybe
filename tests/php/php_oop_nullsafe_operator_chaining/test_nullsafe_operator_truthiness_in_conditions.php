<?php
// vybe-test: php/php_oop_nullsafe_operator_chaining/test_nullsafe_operator_truthiness_in_conditions
// origin: languages/php/tests/php/test_php_oop_nullsafe_operator_chaining.rs

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
    public function score(): int { return 0; }
}
class Container {
    public ?Node $node = null;
}

$ready = new Container();
$ready->node = new Node();
echo $ready->node?->score() ? 'truthy' : 'falsey';
echo '|';
$empty = new Container();
echo $empty->node?->score() ? 'truthy' : 'falsey';
echo '|';
echo (($empty->node?->score() ?: 'fallback'));

__vybe_check(ob_get_clean(), "falsey|falsey|fallback");
