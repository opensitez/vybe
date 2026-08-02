<?php
// vybe-test: php/oop_advanced/object_clone_shallow
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

class Config {
    public string $env = "prod";
    public int $timeout = 30;
}
$a = new Config();
$b = clone $a;
$b->env = "dev";
echo $a->env, "\n";
echo $b->env, "\n";
echo $a->timeout, "\n";

__vybe_check(ob_get_clean(), "prod\ndev\n30");
