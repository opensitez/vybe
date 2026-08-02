<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_magic_call_with_variadics
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

class Handler {
    public function __call(string $name, array $args): mixed {
        if ($name === 'sum') {
            return array_sum($args);
        }
        return null;
    }
}
$handler = new Handler();
$method = 'sum';
echo $handler->$method(1, 2, 3);

__vybe_check(ob_get_clean(), "6");
