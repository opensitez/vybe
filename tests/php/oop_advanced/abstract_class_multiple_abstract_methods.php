<?php
// vybe-test: php/oop_advanced/abstract_class_multiple_abstract_methods
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

abstract class Serializer {
    abstract protected function encode(array $data): string;
    abstract protected function decode(string $raw): array;
    public function roundtrip(array $data): array {
        return $this->decode($this->encode($data));
    }
}
class JsonSerializer extends Serializer {
    protected function encode(array $data): string { return json_encode($data); }
    protected function decode(string $raw): array { return json_decode($raw, true); }
}
$s = new JsonSerializer();
$result = $s->roundtrip(["x" => 1, "y" => 2]);
echo $result["x"], "\n";
echo $result["y"], "\n";

__vybe_check(ob_get_clean(), "1\n2");
