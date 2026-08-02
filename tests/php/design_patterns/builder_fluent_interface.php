<?php
// vybe-test: php/design_patterns/builder_fluent_interface
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

class QueryBuilder {
    private string $table = '';
    private array $conditions = [];
    private ?int $limit = null;
    public function from(string $t): static { $this->table = $t; return $this; }
    public function where(string $c): static { $this->conditions[] = $c; return $this; }
    public function limit(int $n): static { $this->limit = $n; return $this; }
    public function build(): string {
        $sql = "SELECT * FROM $this->table";
        if ($this->conditions) $sql .= ' WHERE ' . implode(' AND ', $this->conditions);
        if ($this->limit) $sql .= " LIMIT $this->limit";
        return $sql;
    }
}
echo (new QueryBuilder)->from('users')->where('age>18')->where('active=1')->limit(10)->build();

__vybe_check(ob_get_clean(), "SELECT * FROM users WHERE age>18 AND active=1 LIMIT 10");
