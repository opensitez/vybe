<?php
// vybe-test: php/php81_features/enum_in_match
// origin: languages/php/tests/php/test_php81_features.rs

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

enum Suit { case Hearts; case Diamonds; case Clubs; case Spades; }
$s = Suit::Hearts;
echo match($s) {
    Suit::Hearts, Suit::Diamonds => 'red',
    default => 'black',
};

__vybe_check(ob_get_clean(), "red");
