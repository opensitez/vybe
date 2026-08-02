<?php
// vybe-test: php/patterns/specification_combinable_rules
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

interface Specification {
    public function isSatisfiedBy($candidate): bool;
}
class AndSpec implements Specification {
    public function __construct(private Specification $a, private Specification $b) {}
    public function isSatisfiedBy($c): bool { return $this->a->isSatisfiedBy($c) && $this->b->isSatisfiedBy($c); }
}
class MinAgeSpec implements Specification {
    public function __construct(private int $min) {}
    public function isSatisfiedBy($c): bool { return $c['age'] >= $this->min; }
}
class ActiveSpec implements Specification {
    public function isSatisfiedBy($c): bool { return $c['active'] === true; }
}
$spec = new AndSpec(new MinAgeSpec(18), new ActiveSpec());
$users = [
    ['age' => 25, 'active' => true],
    ['age' => 15, 'active' => true],
    ['age' => 30, 'active' => false],
];
$count = count(array_filter($users, fn($u) => $spec->isSatisfiedBy($u)));
echo $count;

__vybe_check(ob_get_clean(), "1");
