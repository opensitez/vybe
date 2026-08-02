<?php
// vybe-test: php/oop_advanced/interface_extends_multiple
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

interface Serializable2 {
    public function serialize(): string;
}
interface Deserializable {
    public static function deserialize(string $data): static;
}
interface Codec extends Serializable2, Deserializable {}
class JsonRecord implements Codec {
    public array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function serialize(): string { return json_encode($this->data); }
    public static function deserialize(string $data): static {
        return new static(json_decode($data, true));
    }
}
$r = new JsonRecord(["key" => "value"]);
$s = $r->serialize();
$r2 = JsonRecord::deserialize($s);
echo $s, "\n";
echo $r2->data["key"], "\n";

__vybe_check(ob_get_clean(), "{\"key\":\"value\"}\nvalue");
