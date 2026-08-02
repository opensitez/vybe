<?php
// vybe-test: php/oop_patterns/object_set_state_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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

class Payload {
    public string $value = '';
    public static function __set_state(array $state): self {
        $obj = new self();
        $obj->value = strtoupper($state['value']);
        return $obj;
    }
}
$obj = Payload::__set_state(['value' => 'ok']);
echo $obj->value;

__vybe_check(ob_get_clean(), "OK");
