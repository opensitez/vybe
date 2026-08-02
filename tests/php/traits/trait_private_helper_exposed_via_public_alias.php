<?php
// vybe-test: php/traits/trait_private_helper_exposed_via_public_alias
// origin: languages/php/tests/php/test_traits.rs

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

trait Core {
    private function token(): string { return 'tk'; }
}
class Service {
    use Core { token as public get_token; }
}
echo (new Service())->get_token();

__vybe_check(ob_get_clean(), "tk");
