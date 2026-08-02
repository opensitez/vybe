<?php
// vybe-test: php/anonymous_classes_runtime/anonymous_class_json_serializable
// origin: languages/php/tests/php/test_anonymous_classes_runtime.rs

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

$o = new class implements JsonSerializable {
    public function jsonSerialize(): array { return ['a' => 1]; }
};
echo json_encode($o);

__vybe_check(ob_get_clean(), "{\"a\":1}");
