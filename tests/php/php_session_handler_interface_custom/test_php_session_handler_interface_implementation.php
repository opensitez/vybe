<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_handler_interface_implementation
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs

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

class ArraySessionHandler implements SessionHandlerInterface {
    private array $storage = [];

    public function open(string $path, string $name): bool { return true; }
    public function close(): bool { return true; }
    public function read(string $id): string|false { return $this->storage[$id] ?? ""; }
    public function write(string $id, string $data): bool { $this->storage[$id] = $data; return true; }
    public function destroy(string $id): bool { unset($this->storage[$id]); return true; }
    public function gc(int $max_lifetime): int|false { return 0; }
}

$handler = new ArraySessionHandler();
$res = session_set_save_handler($handler, true);
echo "Handler Registered: " . ($res ? "YES" : "NO");

__vybe_check(ob_get_clean(), "Handler Registered: YES");
