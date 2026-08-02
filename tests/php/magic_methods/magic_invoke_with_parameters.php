<?php
// vybe-test: php/magic_methods/magic_invoke_with_parameters
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Formatter {
    public function __invoke(string $template, ...$args): string {
        return vsprintf($template, $args);
    }
}
$fmt = new Formatter();
echo $fmt("Hello, %s! You are %d years old.", "Alice", 30);

__vybe_check(ob_get_clean(), "Hello, Alice! You are 30 years old.");
