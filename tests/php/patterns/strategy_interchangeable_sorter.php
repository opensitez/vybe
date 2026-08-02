<?php
// vybe-test: php/patterns/strategy_interchangeable_sorter
// origin: languages/php/tests/php/test_patterns.rs

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

interface SortStrategy {
    public function sort(array $data): array;
}
class AscendingSort implements SortStrategy {
    public function sort(array $data): array { sort($data); return $data; }
}
class DescendingSort implements SortStrategy {
    public function sort(array $data): array { rsort($data); return $data; }
}
class Sorter {
    private $strategy;
    public function __construct(SortStrategy $s) { $this->strategy = $s; }
    public function sort(array $data): array { return $this->strategy->sort($data); }
}
$s = new Sorter(new AscendingSort());
echo implode(',', $s->sort([3, 1, 4, 1, 5]));
$s2 = new Sorter(new DescendingSort());
echo implode(',', $s2->sort([3, 1, 4, 1, 5]));

__vybe_check(ob_get_clean(), "1,1,3,4,55,4,3,1,1");
