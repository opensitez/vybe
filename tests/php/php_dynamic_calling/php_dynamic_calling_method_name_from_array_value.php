<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_method_name_from_array_value
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

class DynResolver {
    public function combine(string $a, string $b): string {
        return $a . '-' . $b;
    }
}

$obj = new DynResolver();
$call = ['name' => 'combine'];
echo $obj->{$call['name']}('x', 'y');

__vybe_check(ob_get_clean(), "x-y");
