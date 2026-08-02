<?php
// vybe-test: php/string_escaping/strip_tags_with_allowlist_keeps_permitted_tags
// origin: languages/php/tests/php/test_string_escaping.rs

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

$html = '<p>ok</p><script>x</script><a href="/u">link</a>';
$allowed = strip_tags($html, '<a>');
echo str_contains($allowed, '<a href') && !str_contains($allowed, '<script>') ? 'allow' : 'fail';

__vybe_check(ob_get_clean(), "allow");
