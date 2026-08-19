<?php
// vybe-test: php/patterns/decorator_wraps_same_interface
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

interface Logger {
    public function log(string $msg): void;
}
class ConsoleLogger implements Logger {
    public function log(string $msg): void { echo $msg; }
}
class TimestampDecorator implements Logger {
    private $inner;
    public function __construct(Logger $l) { $this->inner = $l; }
    public function log(string $msg): void { $this->inner->log('[ts] ' . $msg); }
}
class PrefixDecorator implements Logger {
    private $inner;
    private $prefix;
    public function __construct(Logger $l, string $p) { $this->inner = $l; $this->prefix = $p; }
    public function log(string $msg): void { $this->inner->log($this->prefix . ': ' . $msg); }
}
$log = new PrefixDecorator(new TimestampDecorator(new ConsoleLogger()), 'APP');
$log->log('started');

__vybe_check(ob_get_clean(), "[ts] APP: started");
