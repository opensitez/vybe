<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_debug_info_magic_method_custom_dump
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs

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

class UserAccount {
    public function __construct(
        public string $username,
        private string $passwordHash
    ) {}

    public function __debugInfo(): array {
        return [
            "username" => $this->username,
            "passwordHash" => "********" // Mask sensitive data
        ];
    }
}

$user = new UserAccount("alice", "secret_hash");
$exported = var_export($user->__debugInfo(), return: true);
echo str_contains($exported, "********") ? "PASSWORD_MASKED" : "UNMASKED";

__vybe_check(ob_get_clean(), "PASSWORD_MASKED");
