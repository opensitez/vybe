<?php
// vybe-test: php/tokenizer_parse_error_recovery/token_get_all_parse_error
// origin: languages/php/tests/php/test_tokenizer_parse_error_recovery.rs

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

// token_get_all should silently tokenize even if there is a parse error
$source = '<?php class { public }';
$tokens = token_get_all($source);
$count = 0;
foreach ($tokens as $t) {
    if (is_array($t) && (token_name($t[0]) === 'T_CLASS' || token_name($t[0]) === 'T_PUBLIC')) {
        $count++;
    }
}
echo $count;

__vybe_check(ob_get_clean(), "2");
