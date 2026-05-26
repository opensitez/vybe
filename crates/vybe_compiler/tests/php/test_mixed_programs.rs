use super::helpers::run_prints;

// ── Classic algorithms ────────────────────────────────────────

#[test] fn fizzbuzz() {
    assert_eq!(run_prints(r#"<?php
for ($i = 1; $i <= 15; $i++) {
    if ($i % 15 === 0) echo 'FizzBuzz';
    elseif ($i % 3 === 0) echo 'Fizz';
    elseif ($i % 5 === 0) echo 'Buzz';
    else echo $i;
    if ($i < 15) echo ',';
}
"#), vec!["1,2,Fizz,4,Buzz,Fizz,7,8,Fizz,Buzz,11,Fizz,13,14,FizzBuzz"]);
}
#[test] fn sieve_of_eratosthenes() {
    assert_eq!(run_prints(r#"<?php
function sieve(int $limit): array {
    $composite = array_fill(2, $limit - 1, false);
    for ($i = 2; $i * $i <= $limit; $i++) {
        if (!$composite[$i]) {
            for ($j = $i * $i; $j <= $limit; $j += $i) $composite[$j] = true;
        }
    }
    return array_keys(array_filter($composite, fn($v) => !$v));
}
echo implode(',', sieve(30));
"#), vec!["2,3,5,7,11,13,17,19,23,29"]);
}
#[test] fn roman_numeral_conversion() {
    assert_eq!(run_prints(r#"<?php
function toRoman(int $n): string {
    $map = [1000=>'M',900=>'CM',500=>'D',400=>'CD',100=>'C',90=>'XC',50=>'L',40=>'XL',10=>'X',9=>'IX',5=>'V',4=>'IV',1=>'I'];
    $result = '';
    foreach ($map as $val => $sym) { while ($n >= $val) { $result .= $sym; $n -= $val; } }
    return $result;
}
echo toRoman(2024) . ',' . toRoman(42) . ',' . toRoman(1999);
"#), vec!["MMXXIV,XLII,MCMXCIX"]);
}
#[test] fn caesar_cipher() {
    assert_eq!(run_prints(r#"<?php
function caesar(string $text, int $shift): string {
    return preg_replace_callback('/[a-zA-Z]/', function($m) use ($shift) {
        $base = ctype_upper($m[0]) ? ord('A') : ord('a');
        return chr(($ord = ord($m[0]) - $base + $shift) % 26 >= 0 ? $base + $ord % 26 : $base + ($ord % 26 + 26));
    }, $text);
}
echo caesar('Hello World', 13);
"#), vec!["Uryyb Jbeyq"]);
}
#[test] fn luhn_check() {
    assert_eq!(run_prints(r#"<?php
function luhn(string $num): bool {
    $digits = array_reverse(str_split($num));
    $sum = 0;
    foreach ($digits as $i => $d) {
        $d = (int)$d;
        if ($i % 2 === 1) { $d *= 2; if ($d > 9) $d -= 9; }
        $sum += $d;
    }
    return $sum % 10 === 0;
}
echo luhn('4532015112830366') ? 'valid' : 'invalid';
echo ',';
echo luhn('1234567890123456') ? 'valid' : 'invalid';
"#), vec!["valid,invalid"]);
}

// ── Data processing ───────────────────────────────────────────

#[test] fn word_frequency_count() {
    assert_eq!(run_prints(r#"<?php
$text = 'the quick brown fox jumps over the lazy dog the fox';
$words = explode(' ', $text);
$freq = array_count_values($words);
arsort($freq);
$top = array_slice($freq, 0, 2, true);
echo implode(',', array_map(fn($w,$c) => "$w:$c", array_keys($top), array_values($top)));
"#), vec!["the:3,fox:2"]);
}
#[test] fn matrix_multiplication() {
    assert_eq!(run_prints(r#"<?php
function matmul(array $A, array $B): array {
    $result = [];
    for ($i = 0; $i < count($A); $i++) {
        for ($j = 0; $j < count($B[0]); $j++) {
            $result[$i][$j] = 0;
            for ($k = 0; $k < count($B); $k++) {
                $result[$i][$j] += $A[$i][$k] * $B[$k][$j];
            }
        }
    }
    return $result;
}
$A = [[1,2],[3,4]];
$B = [[5,6],[7,8]];
$C = matmul($A, $B);
echo $C[0][0] . ',' . $C[0][1] . ',' . $C[1][0] . ',' . $C[1][1];
"#), vec!["19,22,43,50"]);
}
#[test] fn run_length_encoding() {
    assert_eq!(run_prints(r#"<?php
function rle(string $s): string {
    $result = '';
    $i = 0;
    while ($i < strlen($s)) {
        $c = $s[$i]; $count = 1;
        while ($i + $count < strlen($s) && $s[$i + $count] === $c) $count++;
        $result .= $count > 1 ? $count . $c : $c;
        $i += $count;
    }
    return $result;
}
echo rle('AAABBBCCDDDDEE');
"#), vec!["3A3B2C4D2E"]);
}

// ── String processing ─────────────────────────────────────────

#[test] fn palindrome_check() {
    assert_eq!(run_prints(r#"<?php
function isPalindrome(string $s): bool {
    $s = strtolower(preg_replace('/[^a-zA-Z0-9]/', '', $s));
    return $s === strrev($s);
}
echo isPalindrome('A man a plan a canal Panama') ? 'yes' : 'no';
echo ',';
echo isPalindrome('hello') ? 'yes' : 'no';
"#), vec!["yes,no"]);
}
#[test] fn anagram_check() {
    assert_eq!(run_prints(r#"<?php
function isAnagram(string $a, string $b): bool {
    $sort = function(string $s): string { $arr = str_split(strtolower($s)); sort($arr); return implode($arr); };
    return $sort($a) === $sort($b);
}
echo isAnagram('listen', 'silent') ? 'yes' : 'no';
echo ',';
echo isAnagram('hello', 'world') ? 'yes' : 'no';
"#), vec!["yes,no"]);
}

// ── Number theory ─────────────────────────────────────────────

#[test] fn gcd_and_lcm() {
    assert_eq!(run_prints(r#"<?php
function gcd(int $a, int $b): int { return $b === 0 ? $a : gcd($b, $a % $b); }
function lcm(int $a, int $b): int { return intdiv($a * $b, gcd($a, $b)); }
echo gcd(48, 18) . ',' . lcm(4, 6);
"#), vec!["6,12"]);
}
#[test] fn power_set() {
    assert_eq!(run_prints(r#"<?php
function powerSet(array $set): array {
    if (!$set) return [[]];
    $first = array_shift($set);
    $rest = powerSet($set);
    return array_merge($rest, array_map(fn($s) => array_merge([$first], $s), $rest));
}
$ps = powerSet([1,2,3]);
echo count($ps);
"#), vec!["8"]);
}

// ── Object graph ──────────────────────────────────────────────

#[test] fn linked_list_traversal() {
    assert_eq!(run_prints(r#"<?php
class Node { public ?Node $next = null; public function __construct(public int $val) {} }
$head = new Node(1);
$head->next = new Node(2);
$head->next->next = new Node(3);
$result = [];
for ($n = $head; $n !== null; $n = $n->next) $result[] = $n->val;
echo implode(',', $result);
"#), vec!["1,2,3"]);
}
