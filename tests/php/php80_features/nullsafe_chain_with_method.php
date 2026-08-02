<?php
// vybe-test: php/php80_features/nullsafe_chain_with_method
// origin: languages/php/tests/php/test_php80_features.rs

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

class Session { private ?User2 $user = null; public function getUser(): ?User2 { return $this->user; } }
class User2 { public function getAddress(): ?Addr { return null; } }
class Addr { public function getCity(): string { return 'Paris'; } }
$session = new Session;
echo $session->getUser()?->getAddress()?->getCity() ?? 'no city';

__vybe_check(ob_get_clean(), "no city");
