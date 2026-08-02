<?php
// vybe-test: php/design_patterns/proxy_pattern_lazy_load
// origin: languages/php/tests/php/test_design_patterns.rs

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

interface DataStore { public function get(string $key): mixed; }
class RealStore implements DataStore {
    private array $data = ['name' => 'Alice'];
    public function get(string $key): mixed { echo 'fetched,'; return $this->data[$key] ?? null; }
}
class CachedProxy implements DataStore {
    private array $cache = [];
    private DataStore $store;
    public function __construct() { $this->store = new RealStore; }
    public function get(string $key): mixed {
        if (!isset($this->cache[$key])) $this->cache[$key] = $this->store->get($key);
        return $this->cache[$key];
    }
}
$proxy = new CachedProxy;
echo $proxy->get('name') . ',';
echo $proxy->get('name');

__vybe_check(ob_get_clean(), "fetched,Alice,Alice");
