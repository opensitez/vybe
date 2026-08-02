<?php
// vybe-test: php/compact_extract/compact_used_in_function_context
// origin: languages/php/tests/php/test_compact_extract.rs

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

function makeUser(string $name, int $age, string $role): array {
    return compact('name', 'age', 'role');
}
$u = makeUser('Bob', 25, 'admin');
echo $u['name'] . '/' . $u['role'];

__vybe_check(ob_get_clean(), "Bob/admin");
