<?php
// vybe-test: php/intersection_types/abstract_method_with_intersection_parameter
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

interface Displayable { public function display(): string; }
interface Exportable { public function export(): array; }
abstract class Processor {
    abstract public function handle(Displayable&Exportable $obj): string;
}
class ConcreteProcessor extends Processor {
    public function handle(Displayable&Exportable $obj): string {
        return $obj->display() . ':' . count($obj->export());
    }
}
class Widget implements Displayable, Exportable {
    public function display(): string { return "widget"; }
    public function export(): array { return ['a', 'b']; }
}
echo (new ConcreteProcessor())->handle(new Widget());

__vybe_check(ob_get_clean(), "widget:2");
