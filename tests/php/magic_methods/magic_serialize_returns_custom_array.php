<?php
// vybe-test: php/magic_methods/magic_serialize_returns_custom_array
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Token {
    public function __construct(
        private string $value,
        private int $expiry,
        private string $secret = "internal"
    ) {}
    public function __serialize(): array {
        return ["value" => $this->value, "expiry" => $this->expiry];
    }
    public function __unserialize(array $data): void {
        $this->value  = $data["value"];
        $this->expiry = $data["expiry"];
        $this->secret = "restored";
    }
    public function getValue(): string { return $this->value; }
    public function getSecret(): string { return $this->secret; }
}
$t = new Token("abc123", 9999);
$raw = serialize($t);
$t2 = unserialize($raw);
echo $t2->getValue();
echo $t2->getSecret();

__vybe_check(ob_get_clean(), "abc123restored");
