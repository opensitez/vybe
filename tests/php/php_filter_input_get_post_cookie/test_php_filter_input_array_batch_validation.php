<?php
// vybe-test: php/php_filter_input_get_post_cookie/test_php_filter_input_array_batch_validation
// origin: languages/php/tests/php/test_php_filter_input_get_post_cookie.rs

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

$_POST["email"] = "user@example.com";
$_POST["age"] = "30";

$filters = [
    "email" => FILTER_VALIDATE_EMAIL,
    "age" => [
        "filter" => FILTER_VALIDATE_INT,
        "options" => ["min_range" => 18, "max_range" => 65]
    ]
];

$result = filter_input_array(INPUT_POST, $filters);
echo "Email={$result['email']} Age={$result['age']}";

__vybe_check(ob_get_clean(), "Email=user@example.com Age=30");
