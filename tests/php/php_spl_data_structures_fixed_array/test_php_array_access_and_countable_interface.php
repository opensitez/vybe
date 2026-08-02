<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_array_access_and_countable_interface
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs

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

class ConfigBag implements ArrayAccess, Countable {
    private array $data = [];
    public function offsetExists(mixed $offset): bool { return isset($this->data[$offset]); }
    public function offsetGet(mixed $offset): mixed { return $this->data[$offset] ?? null; }
    public function offsetSet(mixed $offset, mixed $value): void { $this->data[$offset] = $value; }
    public function offsetUnset(mixed $offset): void { unset($this->data[$offset]); }
    public function count(): int { return count($this->data); }
}

$bag = new ConfigBag();
$bag["theme"] = "dark";
$bag["lang"] = "en";
echo count($bag) . " | " . $bag["theme"];

__vybe_check(ob_get_clean(), "2 | dark");
