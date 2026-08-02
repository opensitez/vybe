<?php
// vybe-test: php/null_coalescing_patterns/nullsafe_and_ternary_precedence_runtime
// origin: languages/php/tests/php/test_null_coalescing_patterns.rs

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

class Profile {
    public function city(): ?string { return null; }
}
class User { public ?Profile $profile = null; }
$u = new User();
echo ($u?->profile?->city() ?: 'fallback') . '|';
$u->profile = new Profile();
echo ($u?->profile?->city() ?: 'fallback');

__vybe_check(ob_get_clean(), "fallback|fallback");
