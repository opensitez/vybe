<?php
// vybe-test: php/patterns/dependency_injection_constructor
// origin: languages/php/tests/php/test_patterns.rs

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

interface Storage {
    public function write(string $key, string $val): void;
    public function read(string $key): ?string;
}
class MemoryStorage implements Storage {
    private $data = [];
    public function write(string $key, string $val): void { $this->data[$key] = $val; }
    public function read(string $key): ?string { return $this->data[$key] ?? null; }
}
class Cache {
    public function __construct(private Storage $storage) {}
    public function put(string $k, string $v): void { $this->storage->write($k, $v); }
    public function get(string $k): ?string { return $this->storage->read($k); }
}
$cache = new Cache(new MemoryStorage());
$cache->put('key1', 'value1');
echo $cache->get('key1');
echo $cache->get('missing') ?? 'null';

__vybe_check(ob_get_clean(), "value1null");
