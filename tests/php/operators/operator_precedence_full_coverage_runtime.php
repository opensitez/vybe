<?php
// vybe-test: php/operators/operator_precedence_full_coverage_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo (1 + 2 * 3 - 4 / 2) . '|';
echo ((1 + 2) * 3) . '|';
echo (1 + 2 * 3) . '|';
echo (-2 ** 3) . '|';
echo ((-2) ** 3) . '|';
echo (2 ** 3 ** 2) . '|';
echo ((2 ** 3) ** 2) . '|';
echo (1 + 2 << 1) . '|';
echo (1 + (2 << 1)) . '|';
echo (3 + 4 << 2) . '|';
echo (3 + (4 << 2)) . '|';
echo (7 & 3 | 1) . '|';
echo (7 ^ 3 & 1) . '|';
echo (7 | 3 ^ 1) . '|';
echo ('a' . 1 + 2) . '|';
echo ('a' . (1 + 2)) . '|';
echo (1 < 2 && 2 < 3 ? 'T' : 'F') . '|';
echo (1 < 2 || 2 < 1 ? 'T' : 'F') . '|';
echo (false and true || true ? 'T' : 'F') . '|';
echo (true && false || true ? 'T' : 'F') . '|';
echo (true and false || true ? 'T' : 'F') . '|';
$a = true;
$a = true && false;
echo (($a === false) ? 'F0' : 'T0') . '|';
$a = true;
$a = true and false;
echo (($a === true) ? 'T1' : 'F1') . '|';
echo (0 ?: 2 + 3) . '|';
echo (1 ?: 2 + 3) . '|';
echo (1 ? 2 : 3 + 4) . '|';
echo (0 ? 2 : 3 + 4) . '|';
$payload = ['user' => ['name' => null], 'fallback' => ['name' => 'x']];
echo ($payload['user']['name'] ?? $payload['fallback']['name'] ?? 'none') . '|';
echo (($payload['user']['name'] ?? $payload['fallback']['name']) ?? 'none') . '|';
echo ((0 == false) ? 'T' : 'F') . '|';
echo ((0 === false) ? 'T' : 'F') . '|';
echo ((1 == '1') ? 'T' : 'F') . '|';
echo ((1 === '1') ? 'T' : 'F') . '|';
echo ((1 <=> '1') <=> 0) . '|';
echo (false or true xor true && false ? 'T' : 'F');

__vybe_check(ob_get_clean(), "5|9|7|-8|-8|512|64|6|5|28|19|3|6|7|a3|a3|T|T||T|1|F0|T1|5|1|2|7|x|x|T|F|T|F|0|1");
