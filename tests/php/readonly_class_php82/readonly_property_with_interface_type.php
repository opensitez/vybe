<?php
// vybe-test: php/readonly_class_php82/readonly_property_with_interface_type
// origin: languages/php/tests/php/test_readonly_class_php82.rs

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
class Item implements Identifiable {
    public function __construct(private int $itemId) {}
    public function id(): int { return $this->itemId; }
}
class Container {
    public readonly Identifiable $wrapped;
    public function __construct(Identifiable $item) { $this->wrapped = $item; }
}
$c = new Container(new Item(42));
echo $c->wrapped->id();

__vybe_check(ob_get_clean(), "42");
