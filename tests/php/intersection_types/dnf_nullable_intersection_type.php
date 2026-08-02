<?php
// vybe-test: php/intersection_types/dnf_nullable_intersection_type
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

interface Identifiable { public function id(): int; }
interface Named { public function name(): string; }
class User implements Identifiable, Named {
    public function __construct(private int $id, private string $name) {}
    public function id(): int { return $this->id; }
    public function name(): string { return $this->name; }
}
function describe((Identifiable&Named)|null $entity): string {
    if ($entity === null) return 'none';
    return $entity->id() . ':' . $entity->name();
}
echo describe(new User(1, 'Alice')) . ',' . describe(null);

__vybe_check(ob_get_clean(), "1:Alice,none");
