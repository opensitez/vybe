<?php
// vybe-test: php/design_patterns/strategy_pattern
// origin: languages/php/tests/php/test_design_patterns.rs

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

interface SortStrategy { public function sort(array &$data): void; }
class BubbleSort implements SortStrategy {
    public function sort(array &$data): void { sort($data); }
}
class Sorter {
    public function __construct(private SortStrategy $strategy) {}
    public function sort(array $data): array { $this->strategy->sort($data); return $data; }
}
echo implode(',', (new Sorter(new BubbleSort))->sort([3,1,2]));

__vybe_check(ob_get_clean(), "1,2,3");
