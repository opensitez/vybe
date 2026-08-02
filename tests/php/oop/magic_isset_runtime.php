<?php
// vybe-test: php/oop/magic_isset_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class DynamicStore {
    private array $vals = [];
    public function __set(string $name, mixed $value): void { $this->vals[$name] = $value; }
    public function __get(string $name): mixed { return $this->vals[$name] ?? null; }
    public function __isset(string $name): bool { return array_key_exists($name, $this->vals); }
}
$d = new DynamicStore();
$d->alpha = 1;
echo isset($d->alpha) ? 'alpha' : 'noalpha';
echo '|';
echo isset($d->beta) ? 'beta' : 'nobeta';

__vybe_check(ob_get_clean(), "alpha|nobeta");
