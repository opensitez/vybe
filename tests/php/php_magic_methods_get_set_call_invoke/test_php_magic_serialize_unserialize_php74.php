<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_serialize_unserialize_php74
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs
// vybe-test-mode: compile

class SessionData {
    public string $user = "Alice";
    public string $token = "secret";

    public function __serialize(): array {
        return ["u" => $this->user];
    }
    public function __unserialize(array $data): void {
        $this->user = $data["u"];
        $this->token = "guest";
    }
}

$s = new SessionData();
$str = serialize($s);
$restored = unserialize($str);
echo $restored->user . " " . $restored->token;
