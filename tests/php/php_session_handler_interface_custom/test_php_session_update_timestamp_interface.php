<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_update_timestamp_interface
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

class TimestampSessionHandler implements SessionHandlerInterface, SessionUpdateTimestampHandlerInterface {
    public function open(string $path, string $name): bool { return true; }
    public function close(): bool { return true; }
    public function read(string $id): string|false { return ""; }
    public function write(string $id, string $data): bool { return true; }
    public function destroy(string $id): bool { return true; }
    public function gc(int $max_lifetime): int|false { return 0; }
    public function validateId(string $id): bool { return true; }
    public function updateTimestamp(string $id, string $data): bool { return true; }
}

$handler = new TimestampSessionHandler();
session_set_save_handler($handler);
echo "TIMESTAMP_HANDLER_OK";
