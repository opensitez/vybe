<?php
// vybe-test: php/php_object_mangled_vars_inspection/test_get_mangled_object_vars_public_private_protected
// origin: languages/php/tests/php/test_php_object_mangled_vars_inspection.rs

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

class Sample {
    public int $pub = 1;
    protected string $prot = "secret";
    private bool $priv = true;
}

$vars = get_mangled_object_vars(new Sample());
echo count($vars), "\n";

__vybe_check(ob_get_clean(), "3");
