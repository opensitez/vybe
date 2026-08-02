<?php
// vybe-test: php/classes/class_magic_call_runtime
// origin: languages/php/tests/php/test_classes.rs

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

class MathApi {
    public function __call(string $name, array $args): string {
        return $name . ':' . implode(',', $args);
    }
}
$api = new MathApi();
echo $api->add(2, 3);

__vybe_check(ob_get_clean(), "add:2,3");
