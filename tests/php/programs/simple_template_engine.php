<?php
// vybe-test: php/programs/simple_template_engine
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

function renderTemplate(string $tpl, array $vars): string {
    foreach ($vars as $k => $v) {
        $tpl = str_replace('{{' . $k . '}}', (string)$v, $tpl);
    }
    return $tpl;
}
$tpl = 'Hello, {{name}}! You have {{count}} messages.';
echo renderTemplate($tpl, ['name' => 'Alice', 'count' => 5]) . "\n";

__vybe_check(ob_get_clean(), "Hello, Alice! You have 5 messages.");
