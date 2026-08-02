<?php
// vybe-test: php/oop_advanced/new_self_vs_new_static_in_clone_method
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Base {
    protected string $tag;
    public function __construct(string $tag) { $this->tag = $tag; }
    public function cloneSelf(): self   { return new self("base-copy"); }
    public function cloneStatic(): static { return new static($this->tag . "-copy"); }
    public function getTag(): string { return $this->tag; }
}
class Sub extends Base {}
$s = new Sub("sub");
$a = $s->cloneSelf();
$b = $s->cloneStatic();
echo get_class($a) . ":" . $a->getTag(), "\n";
echo get_class($b) . ":" . $b->getTag(), "\n";

__vybe_check(ob_get_clean(), "Base:base-copy\nSub:sub-copy");
