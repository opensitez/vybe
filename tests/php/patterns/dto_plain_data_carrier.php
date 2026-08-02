<?php
// vybe-test: php/patterns/dto_plain_data_carrier
// origin: languages/php/tests/php/test_patterns.rs

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

class UserDTO {
    public function __construct(
        public readonly int $id,
        public readonly string $name,
        public readonly string $email
    ) {}
    public function toArray(): array {
        return ['id' => $this->id, 'name' => $this->name, 'email' => $this->email];
    }
}
$dto = new UserDTO(1, 'Alice', 'alice@example.com');
echo $dto->name;
echo $dto->toArray()['email'];

__vybe_check(ob_get_clean(), "Alicealice@example.com");
