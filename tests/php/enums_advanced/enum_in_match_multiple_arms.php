<?php
// vybe-test: php/enums_advanced/enum_in_match_multiple_arms
// origin: languages/php/tests/php/test_enums_advanced.rs

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
function color(Suit $s): string {
    return match($s) { Suit::Hearts, Suit::Diamonds => 'red', Suit::Clubs, Suit::Spades => 'black' };
}
echo color(Suit::Diamonds) . ',' . color(Suit::Clubs);

__vybe_check(ob_get_clean(), "red,black");
