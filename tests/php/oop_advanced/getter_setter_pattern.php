<?php
// vybe-test: php/oop_advanced/getter_setter_pattern
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

class Temperature {
    private float $celsius;
    public function __construct(float $celsius) {
        $this->setCelsius($celsius);
    }
    public function getCelsius(): float { return $this->celsius; }
    public function setCelsius(float $val): void { $this->celsius = $val; }
    public function getFahrenheit(): float {
        return $this->celsius * 9 / 5 + 32;
    }
}
$t = new Temperature(100);
echo $t->getCelsius(), "\n";
echo $t->getFahrenheit(), "\n";

__vybe_check(ob_get_clean(), "100\n212");
