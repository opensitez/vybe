<?php
// vybe-test: php/abstract_final_patterns/template_method_pattern_with_abstract
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

abstract class Report {
    final public function generate(): string {
        return $this->header() . "\n" . $this->body() . "\n" . $this->footer();
    }
    abstract protected function header(): string;
    abstract protected function body(): string;
    protected function footer(): string { return "---end---"; }
}
class SalesReport extends Report {
    protected function header(): string { return "SALES REPORT"; }
    protected function body(): string { return "Total: $9999"; }
}
echo (new SalesReport())->generate(), "\n";

__vybe_check(ob_get_clean(), "SALES REPORT\nTotal: \$9999\n---end---");
