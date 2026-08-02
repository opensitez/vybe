<?php
// vybe-test: php/modern_php_deep/nullsafe_combined_with_null_coalescing
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

class Config {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function get(string $key): ?Config {
        return isset($this->data[$key]) ? new Config($this->data[$key]) : null;
    }
    public function value(string $key): mixed { return $this->data[$key] ?? null; }
}
$cfg = new Config(["db" => ["host" => "localhost"]]);
echo $cfg->get("db")?->value("host") ?? "default";
echo $cfg->get("missing")?->value("host") ?? "default";

__vybe_check(ob_get_clean(), "localhostdefault");
