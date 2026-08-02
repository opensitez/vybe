<?php
// vybe-test: php/interfaces_deep/interface_segregation
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface CanRead  { public function read(string $key): mixed; }
interface CanWrite { public function write(string $key, mixed $value): void; }
interface CanDelete { public function delete(string $key): void; }
interface Cache extends CanRead, CanWrite, CanDelete {}
class InMemoryCache implements Cache {
    private array $store = [];
    public function read(string $key): mixed   { return $this->store[$key] ?? null; }
    public function write(string $key, mixed $value): void { $this->store[$key] = $value; }
    public function delete(string $key): void  { unset($this->store[$key]); }
}
function cacheValue(CanWrite $cache, string $key, mixed $value): void {
    $cache->write($key, $value);
}
function readValue(CanRead $cache, string $key): mixed {
    return $cache->read($key);
}
$c = new InMemoryCache();
cacheValue($c, 'name', 'Alice');
echo readValue($c, 'name');
