<?php
// vybe-test: php/oop_advanced/stringable_interface
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

class Money implements Stringable {
    public function __construct(private int $cents) {}
    public function __toString(): string {
        return "$" . number_format($this->cents / 100, 2);
    }
}
function display(Stringable $item): void {
    echo $item, "\n";
}
display(new Money(1299));
display(new Money(50));

__vybe_check(ob_get_clean(), "\$12.99\n\$0.50");
