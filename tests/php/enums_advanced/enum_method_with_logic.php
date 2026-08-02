<?php
// vybe-test: php/enums_advanced/enum_method_with_logic
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

enum HttpMethod: string {
    case GET = 'GET'; case POST = 'POST'; case PUT = 'PUT'; case DELETE = 'DELETE';
    public function isSafe(): bool { return match($this) { self::GET => true, default => false }; }
    public function isIdempotent(): bool { return match($this) { self::GET, self::PUT, self::DELETE => true, default => false }; }
}
echo HttpMethod::GET->isSafe() ? 'safe' : 'unsafe';
echo ',' . (HttpMethod::POST->isIdempotent() ? 'idem' : 'not');

__vybe_check(ob_get_clean(), "safe,not");
