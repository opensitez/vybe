<?php
// vybe-test: php/patterns/builder_fluent_build
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

class Pizza {
    public $size = '';
    public $toppings = [];
    public $crust = '';
}
class PizzaBuilder {
    private $pizza;
    public function __construct() { $this->pizza = new Pizza(); }
    public function size(string $s): self { $this->pizza->size = $s; return $this; }
    public function crust(string $c): self { $this->pizza->crust = $c; return $this; }
    public function topping(string $t): self { $this->pizza->toppings[] = $t; return $this; }
    public function build(): Pizza { return $this->pizza; }
}
$p = (new PizzaBuilder())
    ->size('large')
    ->crust('thin')
    ->topping('mozzarella')
    ->topping('pepperoni')
    ->build();
echo $p->size;
echo $p->crust;
echo implode(',', $p->toppings);

__vybe_check(ob_get_clean(), "largethinmozzarella,pepperoni");
