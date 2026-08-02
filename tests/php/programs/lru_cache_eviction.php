<?php
// vybe-test: php/programs/lru_cache_eviction
// origin: languages/php/tests/php/test_programs.rs

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

class LRUCache {
    private array $cache = [];
    public function __construct(private int $capacity) {}
    public function get(string $key): ?int {
        if (!isset($this->cache[$key])) return null;
        $val = $this->cache[$key];
        unset($this->cache[$key]);
        $this->cache[$key] = $val;
        return $val;
    }
    public function put(string $key, int $val): void {
        if (isset($this->cache[$key])) unset($this->cache[$key]);
        elseif (count($this->cache) >= $this->capacity) array_shift($this->cache);
        $this->cache[$key] = $val;
    }
}
$cache = new LRUCache(3);
$cache->put('a', 1);
$cache->put('b', 2);
$cache->put('c', 3);
echo $cache->get('a') . "\n";
$cache->put('d', 4);
echo ($cache->get('b') === null ? 'null' : $cache->get('b')) . "\n";
echo $cache->get('d') . "\n";

__vybe_check(ob_get_clean(), "1\nnull\n4");
