<?php
// vybe-test: php/oop_advanced/fluent_builder_returns_this
// origin: languages/php/tests/php/test_oop_advanced.rs

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
    private string $table = "";
    private array $conditions = [];
    private ?int $limitVal = null;

    public function from(string $table): static {
        $this->table = $table;
        return $this;
    }
    public function where(string $cond): static {
        $this->conditions[] = $cond;
        return $this;
    }
    public function limit(int $n): static {
        $this->limitVal = $n;
        return $this;
    }
    public function build(): string {
        $sql = "SELECT * FROM {$this->table}";
        if ($this->conditions) {
            $sql .= " WHERE " . implode(" AND ", $this->conditions);
        }
        if ($this->limitVal !== null) {
            $sql .= " LIMIT {$this->limitVal}";
        }
        return $sql;
    }
}
$q = (new QueryBuilder())
    ->from("users")
    ->where("active=1")
    ->where("age>18")
    ->limit(10)
    ->build();
echo $q, "\n";

__vybe_check(ob_get_clean(), "SELECT * FROM users WHERE active=1 AND age>18 LIMIT 10");
