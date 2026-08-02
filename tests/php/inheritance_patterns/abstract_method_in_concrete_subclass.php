<?php
// vybe-test: php/inheritance_patterns/abstract_method_in_concrete_subclass
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

abstract class Serializer {
    abstract protected function encode(mixed $data): string;
    public function serialize(mixed $data): string { return $this->encode($data); }
}
class JsonSerializer extends Serializer {
    protected function encode(mixed $data): string { return json_encode($data); }
}
echo (new JsonSerializer)->serialize(['key' => 'val']), "\n";

__vybe_check(ob_get_clean(), "{\"key\":\"val\"}");
