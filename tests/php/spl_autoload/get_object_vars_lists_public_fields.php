<?php
// vybe-test: php/spl_autoload/get_object_vars_lists_public_fields
// origin: languages/php/tests/php/test_spl_autoload.rs

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

class Holder {
    public $a = 1;
    private $b = 2;
    protected $c = 3;
}
$h = new Holder();
$vars = get_object_vars($h);
echo array_key_exists('a', $vars) ? 'a' : 'na';
echo array_key_exists('b', $vars) ? '|b' : '|nb';
echo array_key_exists('c', $vars) ? '|c' : '|nc';

__vybe_check(ob_get_clean(), "a|nb|nc");
