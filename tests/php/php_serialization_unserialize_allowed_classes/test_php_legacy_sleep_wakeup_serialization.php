<?php
// vybe-test: php/php_serialization_unserialize_allowed_classes/test_php_legacy_sleep_wakeup_serialization
// origin: languages/php/tests/php/test_php_serialization_unserialize_allowed_classes.rs
// vybe-test-mode: compile

class DbModel {
    public string $table = "users";
    public mixed $connection = "active_res";

    public function __sleep(): array {
        return ["table"];
    }

    public function __wakeup(): void {
        $this->connection = "reconnected";
    }
}

$m = new DbModel();
$s = serialize($m);
$restored = unserialize($s);
echo $restored->connection;
