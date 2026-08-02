<?php
// vybe-test: php/patterns/composite_tree_sum
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

interface Component {
    public function price(): int;
}
class Leaf implements Component {
    private $cost;
    public function __construct(int $cost) { $this->cost = $cost; }
    public function price(): int { return $this->cost; }
}
class Composite implements Component {
    private $children = [];
    public function add(Component $c): void { $this->children[] = $c; }
    public function price(): int {
        return array_sum(array_map(fn($c) => $c->price(), $this->children));
    }
}
$box = new Composite();
$box->add(new Leaf(10));
$inner = new Composite();
$inner->add(new Leaf(5));
$inner->add(new Leaf(15));
$box->add($inner);
echo $box->price();

__vybe_check(ob_get_clean(), "30");
