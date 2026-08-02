<?php
// vybe-test: php/generators_advanced/generator_state_machine_lexer
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function tokenize(string $input) {
    $len = strlen($input);
    $i = 0;
    while ($i < $len) {
        if (ctype_space($input[$i])) { $i++; continue; }
        if (ctype_digit($input[$i])) {
            $num = "";
            while ($i < $len && ctype_digit($input[$i])) {
                $num .= $input[$i++];
            }
            yield ["NUM", $num];
        } elseif (str_contains("+-*/", $input[$i])) {
            yield ["OP", $input[$i]];
            $i++;
        } else {
            yield ["UNK", $input[$i]];
            $i++;
        }
    }
}
$tokens = [];
foreach (tokenize("12 + 34 * 5") as [$type, $val]) {
    $tokens[] = "$type:$val";
}
echo implode("|", $tokens);

__vybe_check(ob_get_clean(), "NUM:12|OP:+|NUM:34|OP:*|NUM:5");
