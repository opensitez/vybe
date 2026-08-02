<?php
// vybe-test: php/traits_advanced/multiple_traits_composed
// origin: languages/php/tests/php/test_traits_advanced.rs

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

trait Serializable2 { public function serialize(): string { return json_encode((array)$this); } }
trait Loggable { public function log(): void { echo 'log:' . get_class($this); } }
class Item { use Serializable2, Loggable; public function __construct(public string $name) {} }
$item = new Item('test');
$item->log();
echo ',' . json_decode($item->serialize())->name;

__vybe_check(ob_get_clean(), "log:Item,test");
