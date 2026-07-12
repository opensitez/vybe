use super::helpers::compile_ok;

// ══════════════════════════════════════════════════════════════
// Cross-language compatible types (same host as VB/C#/JS/Python)
// ══════════════════════════════════════════════════════════════

// ── StringBuilder (same as VB/C# System.Text.StringBuilder) ──
#[test]
fn stringbuilder() {
    compile_ok(
        r#"<?php
$sb = new StringBuilder('Hello');
$sb->append(' World');
$sb->appendLine('!');
$sb->insert(5, ',');
echo $sb->toString();
"#,
    );
}

#[test]
fn stringbuilder_replace() {
    compile_ok(
        r#"<?php
$sb = new StringBuilder('Hello World');
$sb->replace('World', 'PHP');
echo $sb->toString();
$sb->clear();
"#,
    );
}

// ── HashSet (same as VB/C# HashSet(Of T)) ───────────────────
#[test]
fn hashset() {
    compile_ok(
        r#"<?php
$set = new HashSet();
$set->add('apple');
$set->add('banana');
$set->add('apple'); // duplicate
echo $set->contains('banana');
$set->remove('banana');
"#,
    );
}

// ── Dictionary (same as VB/C# Dictionary(Of K,V)) ──────────
#[test]
fn dictionary() {
    compile_ok(
        r#"<?php
$dict = new Dictionary();
"#,
    );
}

// ── Random (same as VB/C# System.Random) ────────────────────
#[test]
fn random_obj() {
    compile_ok(
        r#"<?php
$rng = new Random();
$n = $rng->nextInt(1, 100);
$f = $rng->nextFloat();
"#,
    );
}

// ── Stopwatch (same as VB/C# System.Diagnostics.Stopwatch) ──
#[test]
fn stopwatch() {
    compile_ok(
        r#"<?php
$sw = new Stopwatch();
$elapsed = $sw->elapsed();
"#,
    );
}

// ── SplDoublyLinkedList (same as VB/C# LinkedList) ──────────
#[test]
fn linked_list() {
    compile_ok(
        r#"<?php
$list = new SplDoublyLinkedList();
"#,
    );
}

// ── File handles (same as VB Open/Print/Input/Close) ────────
#[test]
fn fopen_fwrite_fclose() {
    compile_ok(
        r#"<?php
$fp = fopen('test.txt', 'w');
fwrite($fp, 'Hello World');
fclose($fp);
"#,
    );
}

#[test]
fn fopen_fgets() {
    compile_ok(
        r#"<?php
$fp = fopen('test.txt', 'r');
$line = fgets($fp);
fclose($fp);
"#,
    );
}

#[test]
fn feof_loop() {
    compile_ok(
        r#"<?php
$fp = fopen('data.csv', 'r');
while (!feof($fp)) {
    $line = fgets($fp);
    echo $line;
}
fclose($fp);
"#,
    );
}

// ══════════════════════════════════════════════════════════════
// Cross-language exception compatibility
// ══════════════════════════════════════════════════════════════

#[test]
fn exception_has_cross_lang_shape() {
    compile_ok(
        r#"<?php
// PHP throw produces same object shape as Python raise, JS throw, VB Throw
try {
    throw new Exception('something failed');
} catch (Exception $e) {
    // $e has __type, __exception_type, name, message — cross-language compatible
    echo $e;
}
"#,
    );
}

#[test]
fn runtime_exception_canonical() {
    compile_ok(
        r#"<?php
try {
    throw new RuntimeException('runtime error');
} catch (RuntimeException $e) {
    // Maps to canonical "RuntimeError" — catchable in Python as RuntimeError
    echo $e;
}
"#,
    );
}

#[test]
fn type_error_canonical() {
    compile_ok(
        r#"<?php
try {
    throw new TypeError('wrong type');
} catch (TypeError $e) {
    // Maps to canonical "TypeError" — catchable in JS as TypeError
    echo $e;
}
"#,
    );
}

// ══════════════════════════════════════════════════════════════
// Cross-language method aliases (automatic via emit_bind_method_with_aliases)
// ══════════════════════════════════════════════════════════════

#[test]
fn method_aliases() {
    compile_ok(
        r#"<?php
// PHP class methods are auto-aliased so other languages can call them:
// toString → __str__ (Python) → ToString (VB/C#)
// contains → __contains__ (Python) → includes (JS)
class Collection {
    public $items = [];
    public function add($item) { array_push($this->items, $item); return $this; }
    public function contains($item) { return in_array($item, $this->items); }
    public function toString() { return implode(', ', $this->items); }
    public function count() { return count($this->items); }
}
$c = new Collection();
$c->add('a')->add('b');
echo $c->toString(); // Also callable as __str__() from Python
echo $c->contains('a'); // Also callable as includes() from JS
echo $c->count(); // Also callable as __len__() from Python
"#,
    );
}

// ══════════════════════════════════════════════════════════════
// Cross-language patterns
// ══════════════════════════════════════════════════════════════

#[test]
fn await_promise() {
    compile_ok(
        r#"<?php
// await() uses same opcode as JS await — can await JS promises
$result = await(fetch('https://api.example.com/data'));
echo $result;
"#,
    );
}

#[test]
fn cross_lang_data_pipeline() {
    compile_ok(
        r#"<?php
// PHP array operations produce same bytecode as Python/JS equivalents
$data = range(1, 10);
$doubled = array_map(fn($n) => $n * 2, $data);
$evens = array_filter($doubled, fn($n) => $n % 4 == 0);
$sum = array_reduce($evens, fn($c, $i) => $c + $i, 0);
echo $sum;
"#,
    );
}

#[test]
fn cross_lang_class_inheritance() {
    compile_ok(
        r#"<?php
// PHP classes use same object layout as Python/JS/VB/C# classes:
// - emit_new_typed_object (same __type stamp)
// - emit_bind_method_with_aliases (cross-lang method names)
// - emit_store_super (__super chain)
// - register_type (type table entry)
// This means a PHP class can extend a Python class at runtime.
class Animal {
    public $name;
    public function __construct($name) { $this->name = $name; }
    public function toString() { return $this->name; }
}
class Dog extends Animal {
    public function speak() { return $this->name . ' barks'; }
}
$d = new Dog('Rex');
echo $d->speak();
echo $d->toString(); // Also callable as __str__() from Python
"#,
    );
}

#[test]
fn component_export() {
    // Original test asserted on per-language compile_component output.
    // Vybex doesn't have a per-language component shim — components are
    // built one layer up. We keep the source-level coverage here.
    compile_ok(
        r#"<?php
function add($a, $b) { return $a + $b; }
function greet($name) { return 'Hello ' . $name; }
class MathHelper {
    public static function square($n) { return $n * $n; }
}
"#,
    );
}
