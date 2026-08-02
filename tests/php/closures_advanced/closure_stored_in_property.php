<?php
// vybe-test: php/closures_advanced/closure_stored_in_property
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Handler {
    private Closure $callback;
    public function __construct(Closure $cb) { $this->callback = $cb; }
    public function handle(string $input): string { return ($this->callback)($input); }
}
$h = new Handler(fn($s) => strtoupper(trim($s)));
echo $h->handle("  hello  ");
