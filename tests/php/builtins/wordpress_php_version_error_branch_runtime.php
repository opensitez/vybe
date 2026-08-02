<?php
// vybe-test: php/builtins/wordpress_php_version_error_branch_runtime
// origin: languages/php/tests/php/test_builtins.rs

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

$required_php_version = '8.2.0'; $wp_version = '6.5.5'; $php_version = '7.0.0'; if ( version_compare( $required_php_version, $php_version, '>' ) ) { printf('Your server is running PHP version %1$s but WordPress %2$s requires at least %3$s.', $php_version, $wp_version, $required_php_version); exit( 1 ); } echo 'after';

__vybe_check(ob_get_clean(), "Your server is running PHP version 7.0.0 but WordPress 6.5.5 requires at least 8.2.0.");
