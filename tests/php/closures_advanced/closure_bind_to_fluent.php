<?php
// vybe-test: php/closures_advanced/closure_bind_to_fluent
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Builder {
    private array $parts = [];
    public function add(string $part): static { $this->parts[] = $part; return $this; }
    public function build(): string { return implode('-', $this->parts); }
}
$addPart = (function(string $p) { $this->parts[] = $p; return $this; })->bindTo(new Builder(), Builder::class);
$b = new Builder();
$b->add('a')->add('b');
echo $b->build();
