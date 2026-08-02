<?php
// vybe-test: php/advanced_oop/template_method_pattern
// origin: languages/php/tests/php/test_advanced_oop.rs

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

abstract class DataProcessor {
    final public function process(array $data): array {
        $data = $this->filter($data);
        $data = $this->transform($data);
        return $this->sort($data);
    }
    abstract protected function filter(array $d): array;
    abstract protected function transform(array $d): array;
    protected function sort(array $d): array { sort($d); return $d; }
}
class EvenDoubler extends DataProcessor {
    protected function filter(array $d): array { return array_filter($d, fn($n) => $n % 2 === 0); }
    protected function transform(array $d): array { return array_map(fn($n) => $n * 2, $d); }
}
echo implode(',', (new EvenDoubler)->process([1,2,3,4,5,6]));

__vybe_check(ob_get_clean(), "4,8,12");
