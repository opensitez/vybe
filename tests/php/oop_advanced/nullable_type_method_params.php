<?php
// vybe-test: php/oop_advanced/nullable_type_method_params
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

class User {
    public function __construct(
        public string $name,
        public ?string $email = null,
    ) {}
    public function contact(): string {
        return $this->email ?? "no email";
    }
}
$u1 = new User("Alice", "alice@example.com");
$u2 = new User("Bob");
echo $u1->contact(), "\n";
echo $u2->contact(), "\n";

__vybe_check(ob_get_clean(), "alice@example.com\nno email");
