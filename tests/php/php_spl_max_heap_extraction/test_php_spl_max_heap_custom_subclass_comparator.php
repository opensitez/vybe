<?php
// vybe-test: php/php_spl_max_heap_extraction/test_php_spl_max_heap_custom_subclass_comparator
// origin: languages/php/tests/php/test_php_spl_max_heap_extraction.rs

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

class ScoreHeap extends SplMaxHeap {
    protected function compare(mixed $val1, mixed $val2): int {
        return $val1["score"] <=> $val2["score"];
    }
}

$heap = new ScoreHeap();
$heap->insert(["player" => "Bob", "score" => 100]);
$heap->insert(["player" => "Alice", "score" => 500]);
$heap->insert(["player" => "Charlie", "score" => 250]);

$top = $heap->extract();
echo "Winner: " . $top["player"] . " (" . $top["score"] . ")";

__vybe_check(ob_get_clean(), "Winner: Alice (500)");
