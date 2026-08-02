<?php
// vybe-test: php/covariant_return_types/chained_covariant_static_returns
// origin: languages/php/tests/php/test_covariant_return_types.rs

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
    protected array $conditions = [];
    public function where(string $cond): static { $this->conditions[] = $cond; return $this; }
    public function sql(): string { return implode(' AND ', $this->conditions); }
}
class UserQuery extends Query {
    public function active(): static { return $this->where('active = 1'); }
}
echo (new UserQuery())->active()->where('age > 18')->sql();

__vybe_check(ob_get_clean(), "active = 1 AND age > 18");
