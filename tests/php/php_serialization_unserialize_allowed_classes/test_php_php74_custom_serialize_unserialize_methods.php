<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_php74_custom_serialize_unserialize_methods
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs

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

class UserRecord {
    public function __construct(
        public int $id,
        public string $username,
        public string $passwordHash
    ) {}

    public function __serialize(): array {
        return ["i" => $this->id, "u" => $this->username];
    }

    public function __unserialize(array $data): void {
        $this->id = $data["i"];
        $this->username = $data["u"];
        $this->passwordHash = "";
    }
}

$u = new UserRecord(1, "john_doe", "secret_hash");
$s = serialize($u);
$restored = unserialize($s);
echo "{$restored->id}:{$restored->username} hash=" . strlen($restored->passwordHash);

__vybe_check(ob_get_clean(), "1:john_doe hash=0");
