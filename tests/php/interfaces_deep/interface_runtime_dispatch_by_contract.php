<?php
// vybe-test: php/interfaces_deep/interface_runtime_dispatch_by_contract
// origin: languages/php/tests/php/test_interfaces_deep.rs

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
    public function log(string $message): void;
}
class MemoryLogger implements Logger {
    public array $events = [];
    public function log(string $message): void {
        $this->events[] = $message;
    }
}

function audit(Logger $logger): void {
    $logger->log('start');
    $logger->log('end');
}

$logger = new MemoryLogger();
audit($logger);
echo implode(',', $logger->events);

__vybe_check(ob_get_clean(), "start,end");
