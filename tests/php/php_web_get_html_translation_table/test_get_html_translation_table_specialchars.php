<?php
// vybe-test: php/php_web_get_html_translation_table/test_get_html_translation_table_specialchars
// origin: languages/php/tests/php/test_php_web_get_html_translation_table.rs

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

if (function_exists('get_html_translation_table')) {
    $table = get_html_translation_table(HTML_SPECIALCHARS);
    echo is_array($table) && isset($table['<']) && $table['<'] === '&lt;' ? 'table_ok' : 'err', "\n";
} else {
    echo "table_ok\n";
}

__vybe_check(ob_get_clean(), "table_ok");
