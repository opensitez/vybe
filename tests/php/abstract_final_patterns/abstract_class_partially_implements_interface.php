<?php
// vybe-test: php/abstract_final_patterns/abstract_class_partially_implements_interface
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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
    public function getLog(): array;
}
abstract class BaseLogger implements Logger {
    protected array $entries = [];
    public function getLog(): array { return $this->entries; }
}
class ConsoleLogger extends BaseLogger {
    public function log(string $msg): void { $this->entries[] = $msg; }
}
$l = new ConsoleLogger();
$l->log("hello");
$l->log("world");
echo implode(',', $l->getLog()), "\n";

__vybe_check(ob_get_clean(), "hello,world");
