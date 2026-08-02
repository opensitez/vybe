<?php
// vybe-test: php/traits_advanced/trait_requires_method_from_using_class
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

trait Printable2 {
    abstract protected function content(): string;
    public function print(): void { echo '[' . $this->content() . ']'; }
}
class Article {
    use Printable2;
    public function __construct(private string $text) {}
    protected function content(): string { return $this->text; }
}
(new Article('Hello World'))->print();

__vybe_check(ob_get_clean(), "[Hello World]");
