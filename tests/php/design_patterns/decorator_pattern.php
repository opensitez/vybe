<?php
// vybe-test: php/design_patterns/decorator_pattern
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

interface TextProcessor { public function process(string $t): string; }
class BaseProcessor implements TextProcessor { public function process(string $t): string { return $t; } }
class TrimDecorator implements TextProcessor {
    public function __construct(private TextProcessor $inner) {}
    public function process(string $t): string { return trim($this->inner->process($t)); }
}
class UpperDecorator implements TextProcessor {
    public function __construct(private TextProcessor $inner) {}
    public function process(string $t): string { return strtoupper($this->inner->process($t)); }
}
$proc = new UpperDecorator(new TrimDecorator(new BaseProcessor));
echo $proc->process('  hello world  ');

__vybe_check(ob_get_clean(), "HELLO WORLD");
