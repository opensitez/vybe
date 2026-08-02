<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_id_interface_create_sid
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

class CustomIdSessionHandler implements SessionHandlerInterface, SessionIdInterface {
    public function open(string $path, string $name): bool { return true; }
    public function close(): bool { return true; }
    public function read(string $id): string|false { return ""; }
    public function write(string $id, string $data): bool { return true; }
    public function destroy(string $id): bool { return true; }
    public function gc(int $max_lifetime): int|false { return 0; }
    public function create_sid(): string { return "custom_id_" . bin2hex(random_bytes(4)); }
}

$handler = new CustomIdSessionHandler();
session_set_save_handler($handler);
echo "CUSTOM_ID_HANDLER_OK";
