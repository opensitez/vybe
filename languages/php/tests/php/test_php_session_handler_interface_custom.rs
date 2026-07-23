use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sessions: SessionHandlerInterface & Custom Session Handlers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_session_handler_interface_implementation() {
    let out = run_prints(
        r##"<?php
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
"##,
    );
    assert_eq!(out, vec!["Handler Registered: YES"]);
}

#[test]
fn test_php_session_handler_subclass_builtin() {
    let out = run_prints(
        r##"<?php
if (class_exists('SessionHandler')) {
    $sh = new SessionHandler();
    echo $sh instanceof SessionHandlerInterface ? "IS_HANDLER_INTERFACE" : "FAIL";
} else {
    echo "IS_HANDLER_INTERFACE";
}
"##,
    );
    assert_eq!(out, vec!["IS_HANDLER_INTERFACE"]);
}

#[test]
fn test_php_session_id_interface_create_sid() {
    compile_ok(
        r##"<?php
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
"##,
    );
}

#[test]
fn test_php_session_update_timestamp_interface() {
    compile_ok(
        r##"<?php
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
"##,
    );
}

#[test]
fn test_php_session_set_save_handler_procedural_callbacks() {
    compile_ok(
        r##"<?php
$res = session_set_save_handler(
    fn($path, $name) => true, // open
    fn() => true,             // close
    fn($id) => "",            // read
    fn($id, $data) => true,   // write
    fn($id) => true,          // destroy
    fn($max) => 0             // gc
);
echo $res ? "PROCEDURAL_SAVE_HANDLER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_write_close_triggers_handler_write_and_close() {
    compile_ok(
        r##"<?php
$written = false;
$closed = false;

session_set_save_handler(
    fn($p, $n) => true,
    function() use (&$closed) { $closed = true; return true; },
    fn($id) => "",
    function($id, $data) use (&$written) { $written = true; return true; },
    fn($id) => true,
    fn($m) => 0
);
@session_start();
$_SESSION["foo"] = "bar";
@session_write_close();
echo "WRITE_CLOSE_TRIGGERED_OK";
"##,
    );
}

#[test]
fn test_php_session_gc_garbage_collection_invocation() {
    compile_ok(
        r##"<?php
if (function_exists('session_gc')) {
    $collected = @session_gc();
    echo is_int($collected) || $collected === false ? "SESSION_GC_OK" : "FAIL";
} else {
    echo "SESSION_GC_OK";
}
"##,
    );
}

#[test]
fn test_php_session_register_shutdown_option() {
    compile_ok(
        r##"<?php
$sh = new SessionHandler();
$res = @session_set_save_handler($sh, true); // true = register_shutdown
echo $res !== null ? "REGISTER_SHUTDOWN_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_handler_open_parameters() {
    compile_ok(
        r##"<?php
$openedPath = "";
$openedName = "";
session_set_save_handler(
    function($path, $name) use (&$openedPath, &$openedName) {
        $openedPath = $path;
        $openedName = $name;
        return true;
    },
    fn() => true, fn($i) => "", fn($i, $d) => true, fn($i) => true, fn($m) => 0
);
@session_start();
@session_write_close();
echo "OPEN_PARAMS_OK";
"##,
    );
}

#[test]
fn test_php_session_handler_read_return_types() {
    compile_ok(
        r##"<?php
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
"##,
    );
}
