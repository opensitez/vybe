use super::helpers::{compile_ok, run_prints};

// ── Classic programs ─────────────────────────────────────────────

#[test]
fn fizzbuzz_1_to_15() {
    assert_eq!(run_prints(r#"<?php
for ($i = 1; $i <= 15; $i++) {
    if ($i % 15 === 0) echo "FizzBuzz\n";
    elseif ($i % 3 === 0) echo "Fizz\n";
    elseif ($i % 5 === 0) echo "Buzz\n";
    else echo $i . "\n";
}
"#), vec!["1","2","Fizz","4","Buzz","Fizz","7","8","Fizz","Buzz","11","Fizz","13","14","FizzBuzz"]);
}

#[test]
fn fibonacci_iterative_first_8() {
    assert_eq!(run_prints(r#"<?php
$a = 0; $b = 1;
for ($i = 0; $i < 8; $i++) {
    echo $a . "\n";
    [$a, $b] = [$b, $a + $b];
}
"#), vec!["0","1","1","2","3","5","8","13"]);
}

#[test]
fn fibonacci_recursive_first_6() {
    assert_eq!(run_prints(r#"<?php
function fib(int $n): int {
    if ($n <= 1) return $n;
    return fib($n - 1) + fib($n - 2);
}
for ($i = 0; $i < 6; $i++) echo fib($i) . "\n";
"#), vec!["0","1","1","2","3","5"]);
}

#[test]
fn factorial_recursive_five() {
    assert_eq!(run_prints(r#"<?php
function fact(int $n): int { return $n <= 1 ? 1 : $n * fact($n - 1); }
echo fact(5) . "\n";
"#), vec!["120"]);
}

#[test]
fn factorial_iterative_seven() {
    assert_eq!(run_prints(r#"<?php
function factIter(int $n): int {
    $r = 1;
    for ($i = 2; $i <= $n; $i++) $r *= $i;
    return $r;
}
echo factIter(7) . "\n";
"#), vec!["5040"]);
}

#[test]
fn bubble_sort_correctness() {
    assert_eq!(run_prints(r#"<?php
function bubbleSort(array $arr): array {
    $n = count($arr);
    for ($i = 0; $i < $n; $i++)
        for ($j = 0; $j < $n - $i - 1; $j++)
            if ($arr[$j] > $arr[$j+1]) { $t=$arr[$j]; $arr[$j]=$arr[$j+1]; $arr[$j+1]=$t; }
    return $arr;
}
echo implode(',', bubbleSort([5,3,1,4,2])) . "\n";
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn insertion_sort_correctness() {
    assert_eq!(run_prints(r#"<?php
function insertionSort(array $arr): array {
    for ($i = 1; $i < count($arr); $i++) {
        $key = $arr[$i];
        $j = $i - 1;
        while ($j >= 0 && $arr[$j] > $key) { $arr[$j+1] = $arr[$j]; $j--; }
        $arr[$j+1] = $key;
    }
    return $arr;
}
echo implode(',', insertionSort([9,3,7,1,5])) . "\n";
"#), vec!["1,3,5,7,9"]);
}

#[test]
fn selection_sort_correctness() {
    assert_eq!(run_prints(r#"<?php
function selectionSort(array $arr): array {
    $n = count($arr);
    for ($i = 0; $i < $n - 1; $i++) {
        $minIdx = $i;
        for ($j = $i+1; $j < $n; $j++) if ($arr[$j] < $arr[$minIdx]) $minIdx = $j;
        [$arr[$i], $arr[$minIdx]] = [$arr[$minIdx], $arr[$i]];
    }
    return $arr;
}
echo implode(',', selectionSort([4,2,6,1,3])) . "\n";
"#), vec!["1,2,3,4,6"]);
}

#[test]
fn binary_search_find_index() {
    assert_eq!(run_prints(r#"<?php
function binarySearch(array $arr, int $target): int {
    $lo = 0; $hi = count($arr) - 1;
    while ($lo <= $hi) {
        $mid = intdiv($lo + $hi, 2);
        if ($arr[$mid] === $target) return $mid;
        elseif ($arr[$mid] < $target) $lo = $mid + 1;
        else $hi = $mid - 1;
    }
    return -1;
}
$sorted = [1,3,5,7,9,11,13];
echo binarySearch($sorted, 7) . "\n";
echo binarySearch($sorted, 1) . "\n";
echo binarySearch($sorted, 6) . "\n";
"#), vec!["3","0","-1"]);
}

#[test]
fn stack_with_array_push_pop_peek() {
    assert_eq!(run_prints(r#"<?php
class Stack {
    private array $data = [];
    public function push($v): void { $this->data[] = $v; }
    public function pop() { return array_pop($this->data); }
    public function peek() { return end($this->data); }
    public function isEmpty(): bool { return empty($this->data); }
    public function size(): int { return count($this->data); }
}
$s = new Stack();
$s->push(1); $s->push(2); $s->push(3);
echo $s->peek() . "\n";
echo $s->pop() . "\n";
echo $s->size() . "\n";
"#), vec!["3","3","2"]);
}

#[test]
fn queue_with_array_enqueue_dequeue() {
    assert_eq!(run_prints(r#"<?php
class Queue {
    private array $data = [];
    public function enqueue($v): void { $this->data[] = $v; }
    public function dequeue() { return array_shift($this->data); }
    public function front() { return $this->data[0] ?? null; }
    public function size(): int { return count($this->data); }
}
$q = new Queue();
$q->enqueue('a'); $q->enqueue('b'); $q->enqueue('c');
echo $q->dequeue() . "\n";
echo $q->front() . "\n";
echo $q->size() . "\n";
"#), vec!["a","b","2"]);
}

#[test]
fn linked_list_append_traverse() {
    assert_eq!(run_prints(r#"<?php
class Node {
    public $next = null;
    public function __construct(public $value) {}
}
class LinkedList {
    private $head = null;
    public function append($v): void {
        $node = new Node($v);
        if ($this->head === null) { $this->head = $node; return; }
        $cur = $this->head;
        while ($cur->next !== null) $cur = $cur->next;
        $cur->next = $node;
    }
    public function toArray(): array {
        $res = []; $cur = $this->head;
        while ($cur !== null) { $res[] = $cur->value; $cur = $cur->next; }
        return $res;
    }
}
$l = new LinkedList();
$l->append(10); $l->append(20); $l->append(30);
echo implode('->', $l->toArray()) . "\n";
"#), vec!["10->20->30"]);
}

#[test]
fn binary_tree_insert_inorder() {
    assert_eq!(run_prints(r#"<?php
class BSTNode {
    public $left = null; public $right = null;
    public function __construct(public int $val) {}
}
class BST {
    private $root = null;
    public function insert(int $v): void { $this->root = $this->insertNode($this->root, $v); }
    private function insertNode(?BSTNode $node, int $v): BSTNode {
        if ($node === null) return new BSTNode($v);
        if ($v < $node->val) $node->left = $this->insertNode($node->left, $v);
        else $node->right = $this->insertNode($node->right, $v);
        return $node;
    }
    public function inorder(): array {
        $result = [];
        $this->traverse($this->root, $result);
        return $result;
    }
    private function traverse(?BSTNode $node, array &$res): void {
        if ($node === null) return;
        $this->traverse($node->left, $res);
        $res[] = $node->val;
        $this->traverse($node->right, $res);
    }
}
$tree = new BST();
foreach ([5,3,7,1,4,6,8] as $v) $tree->insert($v);
echo implode(',', $tree->inorder()) . "\n";
"#), vec!["1,3,4,5,6,7,8"]);
}

#[test]
fn word_frequency_counter() {
    assert_eq!(run_prints(r#"<?php
function wordFrequency(string $text): array {
    $words = str_word_count(strtolower($text), 1);
    $freq = [];
    foreach ($words as $w) $freq[$w] = ($freq[$w] ?? 0) + 1;
    arsort($freq);
    return $freq;
}
$freq = wordFrequency('the cat sat on the mat the cat');
echo $freq['the'] . "\n";
echo $freq['cat'] . "\n";
echo $freq['sat'] . "\n";
"#), vec!["3","2","1"]);
}

#[test]
fn palindrome_check() {
    assert_eq!(run_prints(r#"<?php
function isPalindrome(string $s): bool {
    $s = strtolower(preg_replace('/[^a-zA-Z0-9]/', '', $s));
    return $s === strrev($s);
}
echo isPalindrome('racecar') ? 'true' : 'false';
echo "\n";
echo isPalindrome('hello') ? 'true' : 'false';
echo "\n";
echo isPalindrome('A man a plan a canal Panama') ? 'true' : 'false';
echo "\n";
"#), vec!["true","false","true"]);
}

#[test]
fn anagram_detector() {
    assert_eq!(run_prints(r#"<?php
function isAnagram(string $a, string $b): bool {
    $sortStr = function(string $s): string {
        $chars = str_split(strtolower($s));
        sort($chars);
        return implode('', $chars);
    };
    return $sortStr($a) === $sortStr($b);
}
echo isAnagram('listen', 'silent') ? 'true' : 'false';
echo "\n";
echo isAnagram('hello', 'world') ? 'true' : 'false';
echo "\n";
"#), vec!["true","false"]);
}

#[test]
fn roman_numeral_to_int() {
    assert_eq!(run_prints(r#"<?php
function romanToInt(string $s): int {
    $map = ['I'=>1,'V'=>5,'X'=>10,'L'=>50,'C'=>100,'D'=>500,'M'=>1000];
    $result = 0;
    $prev = 0;
    foreach (array_reverse(str_split($s)) as $c) {
        $v = $map[$c];
        if ($v < $prev) $result -= $v;
        else $result += $v;
        $prev = $v;
    }
    return $result;
}
echo romanToInt('XIV') . "\n";
echo romanToInt('IX') . "\n";
echo romanToInt('XLII') . "\n";
"#), vec!["14","9","42"]);
}

#[test]
fn int_to_roman_numeral() {
    assert_eq!(run_prints(r#"<?php
function intToRoman(int $num): string {
    $vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
    $syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
    $result = '';
    foreach ($vals as $i => $v) {
        while ($num >= $v) { $result .= $syms[$i]; $num -= $v; }
    }
    return $result;
}
echo intToRoman(42) . "\n";
echo intToRoman(9) . "\n";
echo intToRoman(2024) . "\n";
"#), vec!["XLII","IX","MMXXIV"]);
}

#[test]
fn prime_sieve_to_20() {
    assert_eq!(run_prints(r#"<?php
function sieve(int $n): array {
    $is_prime = array_fill(2, $n - 1, true);
    for ($i = 2; $i * $i <= $n; $i++) {
        if ($is_prime[$i]) {
            for ($j = $i * $i; $j <= $n; $j += $i) $is_prime[$j] = false;
        }
    }
    return array_keys(array_filter($is_prime));
}
echo implode(',', sieve(20)) . "\n";
"#), vec!["2,3,5,7,11,13,17,19"]);
}

#[test]
fn gcd_lcm_calculation() {
    assert_eq!(run_prints(r#"<?php
function gcd(int $a, int $b): int { return $b === 0 ? $a : gcd($b, $a % $b); }
function lcm(int $a, int $b): int { return intdiv($a * $b, gcd($a, $b)); }
echo gcd(12, 8) . "\n";
echo gcd(48, 18) . "\n";
echo lcm(4, 6) . "\n";
echo lcm(7, 5) . "\n";
"#), vec!["4","6","12","35"]);
}

#[test]
fn matrix_multiply_2x2() {
    assert_eq!(run_prints(r#"<?php
function matmul(array $a, array $b): array {
    $n = count($a);
    $res = array_fill(0, $n, array_fill(0, $n, 0));
    for ($i = 0; $i < $n; $i++)
        for ($j = 0; $j < $n; $j++)
            for ($k = 0; $k < $n; $k++)
                $res[$i][$j] += $a[$i][$k] * $b[$k][$j];
    return $res;
}
$a = [[1,2],[3,4]];
$b = [[5,6],[7,8]];
$c = matmul($a, $b);
echo $c[0][0] . ',' . $c[0][1] . "\n";
echo $c[1][0] . ',' . $c[1][1] . "\n";
"#), vec!["19,22","43,50"]);
}

#[test]
fn caesar_cipher_encode_decode() {
    assert_eq!(run_prints(r#"<?php
function caesarEncode(string $text, int $shift): string {
    $result = '';
    foreach (str_split($text) as $c) {
        if (ctype_upper($c)) $result .= chr((ord($c) - 65 + $shift) % 26 + 65);
        elseif (ctype_lower($c)) $result .= chr((ord($c) - 97 + $shift) % 26 + 97);
        else $result .= $c;
    }
    return $result;
}
$encoded = caesarEncode('Hello World', 3);
echo $encoded . "\n";
echo caesarEncode($encoded, 23) . "\n";
"#), vec!["Khoor Zruog","Hello World"]);
}

#[test]
fn string_compression_run_length() {
    assert_eq!(run_prints(r#"<?php
function compress(string $s): string {
    if (strlen($s) === 0) return '';
    $result = '';
    $count = 1;
    for ($i = 1; $i <= strlen($s); $i++) {
        if ($i < strlen($s) && $s[$i] === $s[$i-1]) {
            $count++;
        } else {
            $result .= $s[$i-1] . $count;
            $count = 1;
        }
    }
    return $result;
}
echo compress('aabbc') . "\n";
echo compress('aaabbbccc') . "\n";
echo compress('abcd') . "\n";
"#), vec!["a2b2c1","a3b3c3","a1b1c1d1"]);
}

#[test]
fn balanced_parentheses_check() {
    assert_eq!(run_prints(r#"<?php
function isBalanced(string $s): bool {
    $stack = [];
    $pairs = [')'=>'(', ']'=>'[', '}'=>'{'];
    foreach (str_split($s) as $c) {
        if (in_array($c, ['(','[','{'])) $stack[] = $c;
        elseif (isset($pairs[$c])) {
            if (empty($stack) || array_pop($stack) !== $pairs[$c]) return false;
        }
    }
    return empty($stack);
}
echo isBalanced('{[()]}') ? 'true' : 'false';
echo "\n";
echo isBalanced('([)]') ? 'true' : 'false';
echo "\n";
echo isBalanced('((())') ? 'true' : 'false';
echo "\n";
"#), vec!["true","false","false"]);
}

#[test]
fn flatten_nested_array() {
    assert_eq!(run_prints(r#"<?php
function flatten(array $arr): array {
    $result = [];
    foreach ($arr as $item) {
        if (is_array($item)) $result = array_merge($result, flatten($item));
        else $result[] = $item;
    }
    return $result;
}
$nested = [1, [2, [3, 4]], [5, 6], 7];
echo implode(',', flatten($nested)) . "\n";
"#), vec!["1,2,3,4,5,6,7"]);
}

#[test]
fn csv_parser_basic() {
    assert_eq!(run_prints(r#"<?php
function parseCsv(string $input): array {
    return array_map(fn($line) => explode(',', $line), explode("\n", trim($input)));
}
$csv = "name,age,city\nAlice,30,NYC\nBob,25,LA";
$rows = parseCsv($csv);
echo count($rows) . "\n";
echo $rows[0][0] . "\n";
echo $rows[1][1] . "\n";
echo $rows[2][2] . "\n";
"#), vec!["3","name","30","LA"]);
}

#[test]
fn calculator_rpn_evaluator() {
    assert_eq!(run_prints(r#"<?php
function rpn(string $expr): float {
    $stack = [];
    foreach (explode(' ', $expr) as $token) {
        if (is_numeric($token)) {
            $stack[] = (float)$token;
        } else {
            $b = array_pop($stack);
            $a = array_pop($stack);
            match($token) {
                '+' => $stack[] = $a + $b,
                '-' => $stack[] = $a - $b,
                '*' => $stack[] = $a * $b,
                '/' => $stack[] = $a / $b,
            };
        }
    }
    return array_pop($stack);
}
echo rpn('3 4 +') . "\n";
echo rpn('5 1 2 + 4 * + 3 -') . "\n";
"#), vec!["7","14"]);
}

#[test]
fn memoized_fibonacci() {
    assert_eq!(run_prints(r#"<?php
function memoFib(int $n, array &$memo = []): int {
    if ($n <= 1) return $n;
    if (isset($memo[$n])) return $memo[$n];
    $memo[$n] = memoFib($n - 1, $memo) + memoFib($n - 2, $memo);
    return $memo[$n];
}
$m = [];
echo memoFib(10, $m) . "\n";
echo memoFib(20, $m) . "\n";
"#), vec!["55","6765"]);
}

#[test]
fn run_length_encoding_round_trip() {
    assert_eq!(run_prints(r#"<?php
function rleEncode(string $s): string {
    $out = '';
    $i = 0;
    while ($i < strlen($s)) {
        $c = $s[$i]; $cnt = 1;
        while ($i + $cnt < strlen($s) && $s[$i + $cnt] === $c) $cnt++;
        $out .= $cnt . $c;
        $i += $cnt;
    }
    return $out;
}
function rleDecode(string $s): string {
    $out = '';
    preg_match_all('/(\d+)([a-zA-Z])/', $s, $matches, PREG_SET_ORDER);
    foreach ($matches as $m) $out .= str_repeat($m[2], (int)$m[1]);
    return $out;
}
$encoded = rleEncode('AAABBBCCDDDDEE');
echo $encoded . "\n";
echo rleDecode($encoded) . "\n";
"#), vec!["3A3B2C4D2E","AAABBBCCDDDDEE"]);
}

#[test]
fn rot13_encode_decode() {
    assert_eq!(run_prints(r#"<?php
function rot13(string $s): string { return str_rot13($s); }
$encoded = rot13('Hello World');
echo $encoded . "\n";
echo rot13($encoded) . "\n";
"#), vec!["Uryyb Jbeyq","Hello World"]);
}

#[test]
fn base_conversion_decimal_to_binary() {
    assert_eq!(run_prints(r#"<?php
function toBinary(int $n): string {
    if ($n === 0) return '0';
    $bits = '';
    while ($n > 0) { $bits = ($n % 2) . $bits; $n = intdiv($n, 2); }
    return $bits;
}
echo toBinary(0) . "\n";
echo toBinary(10) . "\n";
echo toBinary(255) . "\n";
"#), vec!["0","1010","11111111"]);
}

#[test]
fn simple_template_engine() {
    assert_eq!(run_prints(r#"<?php
function renderTemplate(string $tpl, array $vars): string {
    foreach ($vars as $k => $v) {
        $tpl = str_replace('{{' . $k . '}}', (string)$v, $tpl);
    }
    return $tpl;
}
$tpl = 'Hello, {{name}}! You have {{count}} messages.';
echo renderTemplate($tpl, ['name' => 'Alice', 'count' => 5]) . "\n";
"#), vec!["Hello, Alice! You have 5 messages."]);
}

#[test]
fn trie_structure_insert_search() {
    assert_eq!(run_prints(r#"<?php
class TrieNode {
    public array $children = [];
    public bool $end = false;
}
class Trie {
    private TrieNode $root;
    public function __construct() { $this->root = new TrieNode(); }
    public function insert(string $word): void {
        $node = $this->root;
        foreach (str_split($word) as $c) {
            if (!isset($node->children[$c])) $node->children[$c] = new TrieNode();
            $node = $node->children[$c];
        }
        $node->end = true;
    }
    public function search(string $word): bool {
        $node = $this->root;
        foreach (str_split($word) as $c) {
            if (!isset($node->children[$c])) return false;
            $node = $node->children[$c];
        }
        return $node->end;
    }
    public function startsWith(string $prefix): bool {
        $node = $this->root;
        foreach (str_split($prefix) as $c) {
            if (!isset($node->children[$c])) return false;
            $node = $node->children[$c];
        }
        return true;
    }
}
$trie = new Trie();
$trie->insert('apple');
$trie->insert('app');
echo $trie->search('app') ? 'true' : 'false';
echo "\n";
echo $trie->search('apple') ? 'true' : 'false';
echo "\n";
echo $trie->search('ap') ? 'true' : 'false';
echo "\n";
echo $trie->startsWith('ap') ? 'true' : 'false';
echo "\n";
"#), vec!["true","true","false","true"]);
}

#[test]
fn lru_cache_eviction() {
    assert_eq!(run_prints(r#"<?php
class LRUCache {
    private array $cache = [];
    public function __construct(private int $capacity) {}
    public function get(string $key): ?int {
        if (!isset($this->cache[$key])) return null;
        $val = $this->cache[$key];
        unset($this->cache[$key]);
        $this->cache[$key] = $val;
        return $val;
    }
    public function put(string $key, int $val): void {
        if (isset($this->cache[$key])) unset($this->cache[$key]);
        elseif (count($this->cache) >= $this->capacity) array_shift($this->cache);
        $this->cache[$key] = $val;
    }
}
$cache = new LRUCache(3);
$cache->put('a', 1);
$cache->put('b', 2);
$cache->put('c', 3);
echo $cache->get('a') . "\n";
$cache->put('d', 4);
echo ($cache->get('b') === null ? 'null' : $cache->get('b')) . "\n";
echo $cache->get('d') . "\n";
"#), vec!["1","null","4"]);
}

#[test]
fn json_path_query_nested() {
    assert_eq!(run_prints(r#"<?php
function jsonPath(array $data, string $path) {
    $keys = explode('.', $path);
    $current = $data;
    foreach ($keys as $key) {
        if (!is_array($current) || !array_key_exists($key, $current)) return null;
        $current = $current[$key];
    }
    return $current;
}
$data = ['user' => ['profile' => ['name' => 'Alice', 'age' => 30], 'role' => 'admin']];
echo jsonPath($data, 'user.profile.name') . "\n";
echo jsonPath($data, 'user.role') . "\n";
echo (jsonPath($data, 'user.missing') === null ? 'null' : 'found') . "\n";
"#), vec!["Alice","admin","null"]);
}

#[test]
fn state_machine_traffic_light() {
    assert_eq!(run_prints(r#"<?php
class TrafficLight {
    private array $states = ['red', 'green', 'yellow'];
    private int $index = 0;
    public function current(): string { return $this->states[$this->index]; }
    public function advance(): void { $this->index = ($this->index + 1) % count($this->states); }
}
$light = new TrafficLight();
for ($i = 0; $i < 6; $i++) {
    echo $light->current() . "\n";
    $light->advance();
}
"#), vec!["red","green","yellow","red","green","yellow"]);
}

#[test]
fn event_system_emit_subscribe() {
    assert_eq!(run_prints(r#"<?php
class Events {
    private array $listeners = [];
    public function on(string $event, callable $fn): void { $this->listeners[$event][] = $fn; }
    public function emit(string $event, ...$args): void {
        foreach ($this->listeners[$event] ?? [] as $fn) $fn(...$args);
    }
}
$events = new Events();
$log = [];
$events->on('tick', function(int $n) use (&$log) { $log[] = "tick:$n"; });
$events->on('tick', function(int $n) use (&$log) { $log[] = "tock:$n"; });
$events->emit('tick', 1);
$events->emit('tick', 2);
foreach ($log as $entry) echo $entry . "\n";
"#), vec!["tick:1","tock:1","tick:2","tock:2"]);
}

#[test]
fn number_to_words_small() {
    assert_eq!(run_prints(r#"<?php
function numberToWords(int $n): string {
    $ones = ['','one','two','three','four','five','six','seven','eight','nine',
             'ten','eleven','twelve','thirteen','fourteen','fifteen','sixteen',
             'seventeen','eighteen','nineteen'];
    $tens = ['','','twenty','thirty','forty','fifty','sixty','seventy','eighty','ninety'];
    if ($n < 20) return $ones[$n];
    if ($n < 100) return $tens[intdiv($n,10)] . ($n%10 ? '-' . $ones[$n%10] : '');
    return $ones[intdiv($n,100)] . ' hundred' . ($n%100 ? ' ' . numberToWords($n%100) : '');
}
echo numberToWords(42) . "\n";
echo numberToWords(7) . "\n";
echo numberToWords(100) . "\n";
"#), vec!["forty-two","seven","one hundred"]);
}

#[test]
fn deep_clone_tree() { compile_ok(r#"<?php
class TreeNode {
    public $left = null;
    public $right = null;
    public function __construct(public int $val) {}
    public function deepClone(): self {
        $clone = new self($this->val);
        $clone->left = $this->left ? $this->left->deepClone() : null;
        $clone->right = $this->right ? $this->right->deepClone() : null;
        return $clone;
    }
}
$root = new TreeNode(1);
$root->left = new TreeNode(2);
$root->right = new TreeNode(3);
$root->left->left = new TreeNode(4);
$cloned = $root->deepClone();
$cloned->left->val = 99;
echo $root->left->val;
echo $cloned->left->val;
"#); }

#[test]
fn tokenizer_arithmetic_simple() {
    assert_eq!(run_prints(r#"<?php
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
"#), vec!["5","1","+","3"]);
}

#[test]
fn roman_clock_hours() {
    assert_eq!(run_prints(r#"<?php
function intToRoman(int $n): string {
    $vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
    $syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
    $r = '';
    foreach ($vals as $i => $v) { while ($n >= $v) { $r .= $syms[$i]; $n -= $v; } }
    return $r;
}
foreach ([1, 4, 8, 12] as $h) echo intToRoman($h) . "\n";
"#), vec!["I","IV","VIII","XII"]);
}
