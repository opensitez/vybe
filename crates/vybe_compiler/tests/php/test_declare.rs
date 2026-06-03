use super::helpers::compile_ok;

// ── declare(strict_types=1) ───────────────────────────────────

#[test]
fn strict_types_basic() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function add(int $a, int $b): int { return $a + $b; }
echo add(2, 3);
"#,
    );
}

#[test]
fn strict_types_float_param() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function area(float $r): float { return M_PI * $r * $r; }
echo round(area(2.0), 4);
"#,
    );
}

#[test]
fn strict_types_string_param() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function shout(string $s): string { return strtoupper($s) . '!'; }
echo shout("hello");
"#,
    );
}

#[test]
fn strict_types_bool_param() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function toggle(bool $flag): bool { return !$flag; }
var_dump(toggle(true));
"#,
    );
}

#[test]
fn strict_types_union_type() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function formatId(int|string $id): string { return "ID:$id"; }
echo formatId(42);
echo formatId("uuid-abc");
"#,
    );
}

#[test]
fn strict_types_nullable() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function findUser(?int $id): ?string {
    if ($id === null) return null;
    return "user_$id";
}
echo findUser(5) ?? 'none';
echo findUser(null) ?? 'none';
"#,
    );
}

#[test]
fn strict_types_return_type_enforcement() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function clamp(int $v, int $lo, int $hi): int {
    return max($lo, min($hi, $v));
}
echo clamp(15, 0, 10) . ',' . clamp(-5, 0, 10);
"#,
    );
}

#[test]
fn strict_types_class_method() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
class Money {
    public function __construct(private int $cents) {}
    public function add(int $cents): static {
        $this->cents += $cents;
        return $this;
    }
    public function format(): string { return '$' . number_format($this->cents / 100, 2); }
}
$m = new Money(100);
echo $m->add(50)->format();
"#,
    );
}

#[test]
fn strict_types_type_error_caught() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function strictInt(int $n): int { return $n * 2; }
try {
    $result = strictInt(3);
    echo $result;
} catch (TypeError $e) {
    echo 'type error: ' . $e->getMessage();
}
"#,
    );
}

#[test]
fn strict_types_with_interface() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
interface Measurable {
    public function measure(): float;
}
class Circle implements Measurable {
    public function __construct(private float $r) {}
    public function measure(): float { return M_PI * $this->r ** 2; }
}
$c = new Circle(3.0);
echo round($c->measure(), 2);
"#,
    );
}

// ── declare(ticks=N) ─────────────────────────────────────────

#[test]
fn declare_ticks() {
    compile_ok(
        r#"<?php
$tick_count = 0;
register_tick_function(function() use (&$tick_count) { $tick_count++; });
declare(ticks=1) {
    for ($i = 0; $i < 5; $i++) { $x = $i * 2; }
}
echo $tick_count > 0 ? 'ticked' : 'no ticks';
"#,
    );
}

#[test]
fn declare_ticks_block_form() {
    compile_ok(
        r#"<?php
declare(ticks=1) {
    $sum = 0;
    for ($i = 1; $i <= 10; $i++) { $sum += $i; }
    echo $sum;
}
"#,
    );
}

// ── declare(encoding) ────────────────────────────────────────

#[test]
fn declare_encoding_utf8() {
    compile_ok(
        r#"<?php
declare(encoding='UTF-8');
$s = "héllo";
echo mb_strlen($s);
"#,
    );
}

// ── Strict types + exceptions ─────────────────────────────────

#[test]
fn strict_types_multifile_interaction() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function divide(float $a, float $b): float {
    if ($b == 0.0) throw new \DivisionByZeroError("Division by zero");
    return $a / $b;
}
echo divide(10.0, 4.0);
"#,
    );
}

#[test]
fn strict_types_variadic() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function sumInts(int ...$nums): int { return array_sum($nums); }
echo sumInts(1, 2, 3, 4, 5);
"#,
    );
}

#[test]
fn strict_types_named_args() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function makeTag(string $tag, string $content, bool $self_close = false): string {
    if ($self_close) return "<$tag />";
    return "<$tag>$content</$tag>";
}
echo makeTag(content: 'Hello', tag: 'p');
"#,
    );
}
