<?php
// vybe-test: php/oop/object_serialize_with_magic_runtime
// origin: languages/php/tests/php/test_oop.rs

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
    public string $value = 'x';
    public function __serialize(): array {
        return ['value' => $this->value];
    }
    public function __unserialize(array $data): void {
        $this->value = $data['value'] . '-u';
    }
}
$t = new Token();
$state = serialize($t);
$u = unserialize($state);
echo $u->value;

__vybe_check(ob_get_clean(), "x-u");
