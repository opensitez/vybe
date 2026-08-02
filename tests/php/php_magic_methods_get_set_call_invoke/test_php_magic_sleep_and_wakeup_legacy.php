<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_sleep_and_wakeup_legacy
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs
// vybe-test-mode: compile

class Connection {
    public string $dsn = "sqlite::memory:";
    public function __sleep(): array {
        return ["dsn"];
    }
    public function __wakeup(): void {
        echo "Reconnected";
    }
}

$c = new Connection();
$s = serialize($c);
$restored = unserialize($s);
