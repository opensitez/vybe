use super::helpers::compile_ok;

// ── Pass by reference ─────────────────────────────────────────

#[test] fn pass_by_reference_basic() {
    compile_ok(r#"<?php
function increment(&$val) { $val++; }
$x = 5;
increment($x);
echo $x; // 6
"#);
}

#[test] fn pass_by_reference_swap() {
    compile_ok(r#"<?php
function swap(&$a, &$b) { $tmp = $a; $a = $b; $b = $tmp; }
$x = "hello"; $y = "world";
swap($x, $y);
echo $x . " " . $y;
"#);
}

#[test] fn pass_by_reference_array() {
    compile_ok(r#"<?php
function addItem(array &$arr, $item) { $arr[] = $item; }
$list = [1, 2];
addItem($list, 3);
echo count($list);
"#);
}

#[test] fn pass_by_reference_nested_array() {
    compile_ok(r#"<?php
function doubleAll(array &$arr) {
    foreach ($arr as &$v) { $v *= 2; }
}
$data = [1, 2, 3, 4];
doubleAll($data);
echo implode(',', $data);
"#);
}

#[test] fn pass_by_reference_string_modify() {
    compile_ok(r#"<?php
function uppercase(&$str) { $str = strtoupper($str); }
$s = "hello";
uppercase($s);
echo $s;
"#);
}

// ── Reference assignment ──────────────────────────────────────

#[test] fn reference_assignment_basic() {
    compile_ok(r#"<?php
$a = 10;
$b = &$a;
$b = 20;
echo $a;
"#);
}

#[test] fn reference_chain() {
    compile_ok(r#"<?php
$a = 1;
$b = &$a;
$c = &$b;
$c = 99;
echo $a;
"#);
}

#[test] fn reference_in_array() {
    compile_ok(r#"<?php
$a = 1;
$arr = [&$a, 2, 3];
$a = 100;
echo $arr[0];
"#);
}

#[test] fn reference_array_element() {
    compile_ok(r#"<?php
$arr = ['x' => 1];
$r = &$arr['x'];
$r = 99;
echo $arr['x'];
"#);
}

#[test] fn unset_reference_keeps_original() {
    compile_ok(r#"<?php
$x = 1;
$y = &$x;
unset($y);
$x = 2;
echo $x;
"#);
}

// ── Return by reference ───────────────────────────────────────

#[test] fn return_by_reference_function() {
    compile_ok(r#"<?php
function &getRef(array &$arr, $key) {
    return $arr[$key];
}
$data = ['a' => 1];
$ref = &getRef($data, 'a');
$ref = 99;
echo $data['a'];
"#);
}

#[test] fn return_by_reference_method() {
    compile_ok(r#"<?php
class Counter {
    private int $val = 0;
    public function &getValue(): int { return $this->val; }
}
$c = new Counter();
$ref = &$c->getValue();
$ref = 42;
echo $c->getValue();
"#);
}

#[test] fn return_by_reference_property() {
    compile_ok(r#"<?php
class Config {
    private array $data = [];
    public function &item(string $key): mixed {
        if (!isset($this->data[$key])) { $this->data[$key] = null; }
        return $this->data[$key];
    }
}
$cfg = new Config();
$ref = &$cfg->item('debug');
$ref = true;
echo $cfg->item('debug') ? 'on' : 'off';
"#);
}

// ── Foreach by reference ──────────────────────────────────────

#[test] fn foreach_by_reference() {
    compile_ok(r#"<?php
$arr = [1, 2, 3];
foreach ($arr as &$v) { $v *= 2; }
unset($v);
echo implode(',', $arr);
"#);
}

#[test] fn foreach_reference_nested() {
    compile_ok(r#"<?php
$matrix = [[1, 2], [3, 4]];
foreach ($matrix as &$row) {
    foreach ($row as &$cell) { $cell += 10; }
    unset($cell);
}
unset($row);
echo $matrix[0][0] . ',' . $matrix[1][1];
"#);
}

#[test] fn foreach_reference_reset() {
    compile_ok(r#"<?php
$items = ['a', 'b', 'c'];
foreach ($items as &$item) { $item = strtoupper($item); }
unset($item);
echo implode('', $items);
"#);
}

// ── Reference in closures ─────────────────────────────────────

#[test] fn closure_use_by_reference() {
    compile_ok(r#"<?php
$counter = 0;
$inc = function() use (&$counter) { $counter++; };
$inc(); $inc(); $inc();
echo $counter;
"#);
}

#[test] fn closure_reference_accumulator() {
    compile_ok(r#"<?php
$total = 0;
$add = function(int $n) use (&$total) { $total += $n; };
array_walk([1, 2, 3, 4, 5], $add);
echo $total;
"#);
}

#[test] fn closure_reference_builder() {
    compile_ok(r#"<?php
$result = [];
$collect = function($v) use (&$result) { $result[] = $v * $v; };
array_map($collect, [1, 2, 3]);
echo implode(',', $result);
"#);
}

// ── Global variable reference ─────────────────────────────────

#[test] fn global_reference() {
    compile_ok(r#"<?php
$globalVal = 0;
function modifyGlobal() {
    global $globalVal;
    $globalVal = 42;
}
modifyGlobal();
echo $globalVal;
"#);
}

#[test] fn global_reference_accumulate() {
    compile_ok(r#"<?php
$sum = 0;
function accumulate(int $n) {
    global $sum;
    $sum += $n;
}
foreach ([10, 20, 30] as $v) { accumulate($v); }
echo $sum;
"#);
}

// ── Reference with static variables ──────────────────────────

#[test] fn static_variable_via_reference() {
    compile_ok(r#"<?php
function counter(): int {
    static $n = 0;
    return ++$n;
}
echo counter() . ',' . counter() . ',' . counter();
"#);
}

#[test] fn static_reference_shared() {
    compile_ok(r#"<?php
function nextId(): int {
    static $id = 0;
    $id++;
    return $id;
}
$a = nextId();
$b = nextId();
$c = nextId();
echo "$a,$b,$c";
"#);
}

// ── Reference counting / identity ────────────────────────────

#[test] fn is_same_reference() {
    compile_ok(r#"<?php
$a = [1, 2, 3];
$b = &$a;
$b[] = 4;
echo count($a);
"#);
}

#[test] fn reference_vs_copy() {
    compile_ok(r#"<?php
$a = [1, 2, 3];
$b = $a;     // copy
$c = &$a;    // reference
$b[] = 99;
$c[] = 88;
echo count($a) . ',' . count($b);
"#);
}
