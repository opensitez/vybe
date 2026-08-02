<?php
// vybe-test: php/tokenizer_get_all_line_numbers/token_get_all_tracks_line_numbers
// origin: languages/php/tests/php/test_tokenizer_get_all_line_numbers.rs

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

$source = "<?php\n\n\$x = 1;\n// comment\n";
$tokens = token_get_all($source);
$lines = [];
foreach ($tokens as $token) {
    if (is_array($token)) {
        $name = token_name($token[0]);
        if ($name === 'T_VARIABLE' || $name === 'T_COMMENT') {
            $lines[] = $name . ':' . $token[2];
        }
    }
}
echo implode(',', $lines);

__vybe_check(ob_get_clean(), "T_VARIABLE:3,T_COMMENT:4");
