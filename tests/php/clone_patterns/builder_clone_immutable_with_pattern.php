<?php
// vybe-test: php/clone_patterns/builder_clone_immutable_with_pattern
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Query {
    private array $filters = [];
    public function where(string $filter): static {
        $new = clone $this;
        $new->filters[] = $filter;
        return $new;
    }
    public function build(): string { return implode(' AND ', $this->filters); }
}
$base = new Query();
$q1 = $base->where('a=1')->where('b=2');
$q2 = $base->where('c=3');
echo $q1->build() . '|' . $q2->build();

__vybe_check(ob_get_clean(), "a=1 AND b=2|c=3");
