<?php
// vybe-test: php/programs/json_path_query_nested
// origin: languages/php/tests/php/test_programs.rs

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

function jsonPath(array $data, string $path) {
    $keys = explode('.', $path);
    $current = $data;
    foreach ($keys as $key) {
        if (!is_array($current) || !array_key_exists($key, $current)) return null;
        $current = $current[$key];
    }
    return $current;
}
$data = ['user' => ['profile' => ['name' => 'Alice', 'age' => 30], 'role' => 'admin']];
echo jsonPath($data, 'user.profile.name') . "\n";
echo jsonPath($data, 'user.role') . "\n";
echo (jsonPath($data, 'user.missing') === null ? 'null' : 'found') . "\n";

__vybe_check(ob_get_clean(), "Alice\nadmin\nnull");
