<?php
// vybe-test: php/programs/tokenizer_arithmetic_simple
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

function tokenize(string $expr): array {
    $tokens = [];
    $i = 0;
    while ($i < strlen($expr)) {
        if (ctype_space($expr[$i])) { $i++; continue; }
        if (ctype_digit($expr[$i])) {
            $num = '';
            while ($i < strlen($expr) && ctype_digit($expr[$i])) { $num .= $expr[$i]; $i++; }
            $tokens[] = ['type' => 'num', 'val' => (int)$num];
        } else {
            $tokens[] = ['type' => 'op', 'val' => $expr[$i]];
            $i++;
        }
    }
    return $tokens;
}
$tokens = tokenize('1 + 2 * 3');
echo count($tokens) . "\n";
echo $tokens[0]['val'] . "\n";
echo $tokens[1]['val'] . "\n";
echo $tokens[4]['val'] . "\n";

__vybe_check(ob_get_clean(), "5\n1\n+\n3");
