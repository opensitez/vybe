<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_handler_read_return_types
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

class StrictReadHandler implements SessionHandlerInterface {
    public function open($p, $n): bool { return true; }
    public function close(): bool { return true; }
    public function read($id): string { return "key|s:3:\"val\";"; }
    public function write($id, $data): bool { return true; }
    public function destroy($id): bool { return true; }
    public function gc($max): int { return 0; }
}
session_set_save_handler(new StrictReadHandler());
@session_start();
echo ($_SESSION["key"] ?? "") === "val" ? "STRICT_READ_OK" : "FAIL";
@session_write_close();
