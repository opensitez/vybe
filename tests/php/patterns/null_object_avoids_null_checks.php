<?php
// vybe-test: php/patterns/null_object_avoids_null_checks
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
class NullLogger implements Logger {
    public function log(string $msg): void {}
}
function processData(array $data, Logger $logger): int {
    $logger->log('processing');
    return count($data);
}
echo processData([1, 2, 3], new ConsoleLogger());
echo processData([1, 2], new NullLogger());

__vybe_check(ob_get_clean(), "processing32");
