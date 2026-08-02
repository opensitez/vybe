<?php
// vybe-test: php/references_advanced/function_returns_reference
// origin: languages/php/tests/php/test_references_advanced.rs

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
    private array $data = ['key' => 'value'];
    public function &get(string $k): mixed { return $this->data[$k]; }
}
$cfg = new Config;
$ref = &$cfg->get('key');
$ref = 'changed';
echo $cfg->get('key');

__vybe_check(ob_get_clean(), "changed");
