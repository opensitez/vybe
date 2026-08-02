<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_match_subject_as_nullsafe_chain
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs

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
    public function tier(): ?string {
        return "pro";
    }
}
class User {
    public ?Profile $profile = null;
}
$u = new User();
$level = match ($u->profile?->tier()) {
    null => 'none',
    'pro' => 'pro-user',
    default => 'other',
};
echo $level;
echo '|';
$u->profile = new Profile();
$level2 = match ($u->profile?->tier()) {
    null => 'none',
    'pro' => 'pro-user',
    default => 'other',
};
echo $level2;

__vybe_check(ob_get_clean(), "none|pro-user");
