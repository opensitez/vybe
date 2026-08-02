<?php
// vybe-test: php/method_chaining/chain_with_static_factory_and_fluent_access
// origin: languages/php/tests/php/test_method_chaining.rs

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

class Service {
    private array $events = [];
    public static function from(string $first): static {
        return new static($first);
    }
    private function __construct(private string $seed) {}
    public function append(string $part): static {
        $this->events[] = $part;
        return $this;
    }
    public function summary(): string {
        return $this->seed . ':' . implode('+', $this->events);
    }
}
echo Service::from('root')->append('a')->append('b')->summary();

__vybe_check(ob_get_clean(), "root:a+b");
