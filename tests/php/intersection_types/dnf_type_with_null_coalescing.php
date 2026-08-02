<?php
// vybe-test: php/intersection_types/dnf_type_with_null_coalescing
// origin: languages/php/tests/php/test_intersection_types.rs

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

interface HasId { public function id(): int; }
interface HasName { public function name(): string; }
class Entity implements HasId, HasName {
    public function __construct(private int $id, private string $name) {}
    public function id(): int { return $this->id; }
    public function name(): string { return $this->name; }
}
function display((HasId&HasName)|null $e): string {
    return $e?->name() ?? 'anonymous';
}
echo display(new Entity(1, 'Alice')) . ',' . display(null);

__vybe_check(ob_get_clean(), "Alice,anonymous");
