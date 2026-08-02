<?php
// vybe-test: php/intersection_types/dnf_type_used_in_function_dispatch
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

interface Sizeable { public function size(): int; }
interface Nameable { public function name(): string; }
class File implements Sizeable, Nameable {
    public function __construct(private string $n, private int $s) {}
    public function size(): int { return $this->s; }
    public function name(): string { return $this->n; }
}
class Unknown {}
function describe((Sizeable&Nameable)|Unknown $item): string {
    if ($item instanceof Unknown) return 'unknown';
    return $item->name() . ':' . $item->size();
}
$items = [new File('a.txt', 100), new Unknown(), new File('b.php', 200)];
$results = array_map('describe', $items);
echo implode(',', $results);

__vybe_check(ob_get_clean(), "a.txt:100,unknown,b.php:200");
