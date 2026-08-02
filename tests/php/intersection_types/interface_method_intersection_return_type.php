<?php
// vybe-test: php/intersection_types/interface_method_intersection_return_type
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

interface Source { public function source(): string; }
interface Sink { public function sink(string $s): void; }
interface Pipe {
    public function getTransformer(): Source&Sink;
}
class PassThrough implements Source, Sink {
    private string $buf = '';
    public function source(): string { return $this->buf; }
    public function sink(string $s): void { $this->buf = strtoupper($s); }
}
class Pipeline implements Pipe {
    private PassThrough $t;
    public function __construct() { $this->t = new PassThrough(); }
    public function getTransformer(): Source&Sink { return $this->t; }
}
$p = new Pipeline();
$t = $p->getTransformer();
$t->sink("hello");
echo $t->source();

__vybe_check(ob_get_clean(), "HELLO");
