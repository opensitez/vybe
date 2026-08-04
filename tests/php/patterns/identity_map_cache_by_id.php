<?php
// vybe-test: php/patterns/identity_map_cache_by_id
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

class IdentityMap {
    private $map = [];
    private $loads = 0;
    public function get(string $type, int $id, callable $loader) {
        $key = "$type:$id";
        if (!isset($this->map[$key])) {
            $this->loads++;
            $this->map[$key] = $loader($id);
        }
        return $this->map[$key];
    }
    public function loads(): int { return $this->loads; }
}
$idmap = new IdentityMap();
$u1 = $idmap->get('user', 1, fn($id) => (object)['id' => $id, 'name' => 'Alice']);
$u2 = $idmap->get('user', 1, fn($id) => (object)['id' => $id, 'name' => 'SHOULDNOTRUN']);
$u3 = $idmap->get('user', 2, fn($id) => (object)['id' => $id, 'name' => 'Bob']);
echo $u1->name;
echo ($u1 === $u2) ? 'cached' : 'dup';
echo $idmap->loads();

__vybe_check(ob_get_clean(), "Alicecached2");
