<?php
// vybe-test: php/php_oop_nullsafe_operator_chaining/test_nullsafe_with_ternary_and_space_operator_style
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

class Score {
    public int $value;
    public function __construct(int $value) { $this->value = $value; }
}
class Box {
    public ?Score $score;
    public function __construct(?Score $score = null) { $this->score = $score; }
}

$with = new Box(new Score(7));
$without = new Box(null);
$left = $with->score?->value;
$right = $without->score?->value;
echo (($left <=> 5) > 0) ? 'gt' : 'lte';
echo '|';
echo (($right <=> 5) > 0) ? 'gt' : (($right ?? 0) ? 'truthy' : 'false');

__vybe_check(ob_get_clean(), "gt|false");
