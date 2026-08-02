<?php
// vybe-test: php/patterns/pipeline_chain_callables
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

class Pipeline {
    private $stages = [];
    public function pipe(callable $fn): self { $this->stages[] = $fn; return $this; }
    public function process($payload) {
        return array_reduce($this->stages, fn($carry, $fn) => $fn($carry), $payload);
    }
}
$result = (new Pipeline())
    ->pipe(fn($s) => trim($s))
    ->pipe(fn($s) => strtoupper($s))
    ->pipe(fn($s) => str_replace(' ', '_', $s))
    ->process('  hello world  ');
echo $result;

__vybe_check(ob_get_clean(), "HELLO_WORLD");
