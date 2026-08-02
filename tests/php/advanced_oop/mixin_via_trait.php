<?php
// vybe-test: php/advanced_oop/mixin_via_trait
// origin: languages/php/tests/php/test_advanced_oop.rs

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

trait HasUuid {
    private string $uuid;
    public function initUuid(): void { $this->uuid = sprintf('%08x-%04x-%04x', 1, 2, 3); }
    public function getUuid(): string { return $this->uuid; }
}
class Entity { use HasUuid; }
$e = new Entity; $e->initUuid();
echo str_contains($e->getUuid(), '-') ? 'has-uuid' : 'no';

__vybe_check(ob_get_clean(), "has-uuid");
