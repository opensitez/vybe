<?php
// vybe-test: php/oop_advanced/abstract_template_method
// origin: languages/php/tests/php/test_oop_advanced.rs

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

abstract class Report {
    abstract protected function getData(): array;
    public function generate(): string {
        $data = $this->getData();
        return implode(", ", $data);
    }
}
class SalesReport extends Report {
    protected function getData(): array {
        return ["Q1: 100", "Q2: 200", "Q3: 150"];
    }
}
$r = new SalesReport();
echo $r->generate(), "\n";

__vybe_check(ob_get_clean(), "Q1: 100, Q2: 200, Q3: 150");
