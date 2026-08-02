<?php
// vybe-test: php/oop_patterns/method_chaining_returns_this
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class QueryBuilder {
    private string $table  = '';
    private array  $wheres = [];
    private ?int   $limit  = null;
    public function from(string $t): static { $this->table = $t; return $this; }
    public function where(string $cond): static { $this->wheres[] = $cond; return $this; }
    public function limit(int $n): static { $this->limit = $n; return $this; }
    public function toSql(): string {
        $sql = 'SELECT * FROM ' . $this->table;
        if ($this->wheres) $sql .= ' WHERE ' . implode(' AND ', $this->wheres);
        if ($this->limit !== null) $sql .= ' LIMIT ' . $this->limit;
        return $sql;
    }
}
$q = (new QueryBuilder())->from('users')->where('active=1')->where('age>18')->limit(10);
echo $q->toSql();
