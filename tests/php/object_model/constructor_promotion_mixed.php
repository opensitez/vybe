<?php
// vybe-test: php/object_model/constructor_promotion_mixed
// origin: languages/php/tests/php/test_object_model.rs

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

class User {
    public string $fullname;
    public function __construct(public string $first, public string $last) {
        $this->fullname = "$first $last";
    }
}
$u = new User('John', 'Doe');
echo $u->fullname;

__vybe_check(ob_get_clean(), "John Doe");
