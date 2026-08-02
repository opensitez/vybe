<?php
// vybe-test: php/magic_methods/magic_clone_with_array_property
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Collection {
    public array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function __clone() {
        $this->items = array_reverse($this->items);
    }
}
$a = new Collection([1, 2, 3]);
$b = clone $a;
$b->items[] = 4;
echo implode(",", $a->items);
echo implode(",", $b->items);

__vybe_check(ob_get_clean(), "1,2,33,2,1,4");
