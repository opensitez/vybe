<?php
// vybe-test: php/oop_advanced/interface_dispatch_with_runtime_check
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

interface Transform {
    public function apply(string $value): string;
}
class JsonTransform implements Transform {
    public function apply(string $value): string {
        return json_encode(["value" => $value]);
    }
}
$transform = new JsonTransform();
echo ($transform instanceof Transform) ? "yes" : "no";
echo "|";
echo $transform->apply("x"), "\n";

__vybe_check(ob_get_clean(), "yes|{\"value\":\"x\"}");
