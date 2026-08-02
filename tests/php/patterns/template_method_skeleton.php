<?php
// vybe-test: php/patterns/template_method_skeleton
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

abstract class DataExporter {
    final public function export(): string {
        $data = $this->fetchData();
        $formatted = $this->format($data);
        return $this->output($formatted);
    }
    abstract protected function fetchData(): array;
    abstract protected function format(array $data): string;
    protected function output(string $s): string { return 'OUT:' . $s; }
}
class CsvExporter extends DataExporter {
    protected function fetchData(): array { return [1, 2, 3]; }
    protected function format(array $data): string { return implode(',', $data); }
}
class JsonExporter extends DataExporter {
    protected function fetchData(): array { return ['a' => 1]; }
    protected function format(array $data): string { return json_encode($data); }
}
echo (new CsvExporter())->export();
echo (new JsonExporter())->export();

__vybe_check(ob_get_clean(), "OUT:1,2,3OUT:{\"a\":1}");
