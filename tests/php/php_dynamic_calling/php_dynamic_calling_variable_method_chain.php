<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_variable_method_chain
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

class ChainNode {
    public function level(int $n): string {
        $method = 'suffix';
        return $this->$method($n);
    }
    public function suffix(int $n): string { return 'v'.$n; }
}
$obj = new ChainNode();
echo $obj->{"level"}(9);

__vybe_check(ob_get_clean(), "v9");
