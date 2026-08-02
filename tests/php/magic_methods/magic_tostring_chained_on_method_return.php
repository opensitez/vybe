<?php
// vybe-test: php/magic_methods/magic_tostring_chained_on_method_return
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

class Tag {
    public function __construct(private string $html) {}
    public function __toString(): string { return $this->html; }
    public function wrap(string $tag): self {
        return new self("<$tag>" . $this->html . "</$tag>");
    }
}
$t = new Tag("hello");
echo $t->wrap("b");

__vybe_check(ob_get_clean(), "<b>hello</b>");
