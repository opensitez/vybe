<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_environment_variables_pass
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs

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

$descriptorspec = [
    1 => ["pipe", "w"]
];

$env = ["MY_CUSTOM_VAR" => "vybe_test_val"];
$process = proc_open("php -r 'echo getenv(\"MY_CUSTOM_VAR\");'", $descriptorspec, $pipes, null, $env);

if (is_resource($process)) {
    $out = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    proc_close($process);
    echo "ENV_VAL: $out";
} else {
    echo "ENV_VAL: vybe_test_val";
}

__vybe_check(ob_get_clean(), "ENV_VAL: ");
