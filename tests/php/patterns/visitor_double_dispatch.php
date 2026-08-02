<?php
// vybe-test: php/patterns/visitor_double_dispatch
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

interface Visitor {
    public function visitCircle(Circle $c): string;
    public function visitRect(Rect $r): string;
}
interface Shape {
    public function accept(Visitor $v): string;
}
class Circle implements Shape {
    public function __construct(public float $r) {}
    public function accept(Visitor $v): string { return $v->visitCircle($this); }
}
class Rect implements Shape {
    public function __construct(public float $w, public float $h) {}
    public function accept(Visitor $v): string { return $v->visitRect($this); }
}
class AreaVisitor implements Visitor {
    public function visitCircle(Circle $c): string { return (string)round(M_PI * $c->r * $c->r, 2); }
    public function visitRect(Rect $r): string { return (string)($r->w * $r->h); }
}
$v = new AreaVisitor();
echo (new Circle(2.0))->accept($v);
echo (new Rect(3.0, 4.0))->accept($v);

__vybe_check(ob_get_clean(), "12.5712");
