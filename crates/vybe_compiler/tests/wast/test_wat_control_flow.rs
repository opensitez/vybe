/// Tests for WAT structured control flow — block, loop, if/else, br, br_if, br_table
use super::helpers::{parse_ok, compile_ok};

// ── block ─────────────────────────────────────────────────────────────────────

#[test]
fn block_empty() {
    parse_ok("(module (func (block)))");
}

#[test]
fn block_with_label() {
    parse_ok("(module (func (block $b nop)))");
}

#[test]
fn block_with_result_type() {
    parse_ok("(module (func (result i32) (block (result i32) i32.const 1)))");
}

#[test]
fn block_br_break() {
    parse_ok("(module (func (block $b i32.const 1 drop br $b)))");
}

#[test]
fn nested_blocks() {
    parse_ok("(module (func (block $outer (block $inner br $outer))))");
}

// ── loop ──────────────────────────────────────────────────────────────────────

#[test]
fn loop_empty() {
    parse_ok("(module (func (loop)))");
}

#[test]
fn loop_with_label() {
    parse_ok("(module (func (loop $l nop)))");
}

#[test]
fn loop_with_br_continue() {
    parse_ok("(module (func (param i32) (local i32) (loop $l local.get 0 local.get 1 i32.add local.set 1 local.get 0 i32.const 1 i32.sub local.tee 0 br_if $l)))");
}

#[test]
fn loop_with_result() {
    parse_ok("(module (func (result i32) (loop (result i32) i32.const 42)))");
}

// ── if / else ─────────────────────────────────────────────────────────────────

#[test]
fn if_no_else_plain() {
    parse_ok("(module (func (param i32) local.get 0 if nop end))");
}

#[test]
fn if_else_plain() {
    parse_ok("(module (func (param i32) (result i32) local.get 0 if (result i32) i32.const 1 else i32.const 0 end))");
}

#[test]
fn if_with_label() {
    parse_ok("(module (func (param i32) local.get 0 if $l nop end))");
}

#[test]
fn if_folded_no_else() {
    parse_ok("(module (func (param i32) (if (local.get 0) (then nop))))");
}

#[test]
fn if_folded_with_else() {
    parse_ok("(module (func (param i32) (result i32) (if (result i32) (local.get 0) (then (i32.const 1)) (else (i32.const 0)))))");
}

#[test]
fn if_folded_nested() {
    parse_ok(r#"
(module
  (func (param i32) (result i32)
    (if (result i32) (local.get 0)
      (then
        (if (result i32) (i32.const 1)
          (then (i32.const 10))
          (else (i32.const 20))))
      (else (i32.const 0)))))
"#);
}

#[test]
fn if_compiles() {
    compile_ok(r#"
(module
  (func $abs (export "abs") (param $x i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (i32.const 0))
      (then (i32.sub (i32.const 0) (local.get $x)))
      (else (local.get $x)))))
"#);
}

// ── br / br_if / br_table ─────────────────────────────────────────────────────

#[test]
fn br_to_block() {
    parse_ok("(module (func (block $b br $b)))");
}

#[test]
fn br_to_loop() {
    parse_ok("(module (func (loop $l br $l)))");
}

#[test]
fn br_if_conditional() {
    parse_ok("(module (func (param i32) (block $b local.get 0 br_if $b)))");
}

#[test]
fn br_table_simple() {
    parse_ok("(module (func (param i32) (block $a (block $b local.get 0 br_table $a $b))))");
}

#[test]
fn br_table_default() {
    parse_ok("(module (func (param i32) (block $a (block $b (block $c local.get 0 br_table $a $b $c)))))");
}

// ── return ────────────────────────────────────────────────────────────────────

#[test]
fn return_void() {
    parse_ok("(module (func return))");
}

#[test]
fn return_value() {
    parse_ok("(module (func (result i32) i32.const 42 return))");
}

#[test]
fn early_return_in_if() {
    parse_ok(r#"
(module
  (func (export "f") (param i32) (result i32)
    local.get 0
    if (result i32)
      i32.const 1
      return
    end
    i32.const 0))
"#);
}

// ── return_call / return_call_indirect (tail calls) ───────────────────────────

#[test]
fn return_call() {
    parse_ok("(module (func $f (param i32) (result i32) local.get 0) (func (param i32) (result i32) local.get 0 return_call $f))");
}

#[test]
fn return_call_indirect() {
    parse_ok("(module (type $t (func (param i32) (result i32))) (table 1 funcref) (func (param i32 i32) (result i32) local.get 0 local.get 1 return_call_indirect (type $t)))");
}

// ── Exceptions proposal ───────────────────────────────────────────────────────

#[test]
fn try_catch_folded() {
    parse_ok(r#"
(module
  (tag $e (param i32))
  (func (export "f")
    (try
      (nop)
      (catch $e drop))))
"#);
}

#[test]
fn throw_instr() {
    parse_ok(r#"
(module
  (tag $e (param i32))
  (func (export "f") i32.const 42 throw $e))
"#);
}

#[test]
fn rethrow_instr() {
    parse_ok(r#"
(module
  (tag $e)
  (func (export "f")
    try
      nop
    catch $e
      rethrow 0
    end))
"#);
}

// ── Compile checks ────────────────────────────────────────────────────────────

#[test]
fn compile_loop_sum() {
    compile_ok(r#"
(module
  (func $sum (export "sum") (param $n i32) (result i32)
    (local $acc i32)
    (local $i i32)
    i32.const 0
    local.set $acc
    i32.const 1
    local.set $i
    (block $break
      (loop $continue
        local.get $i
        local.get $n
        i32.gt_s
        br_if $break
        local.get $acc
        local.get $i
        i32.add
        local.set $acc
        local.get $i
        i32.const 1
        i32.add
        local.set $i
        br $continue))
    local.get $acc))
"#);
}

#[test]
fn compile_max_func() {
    compile_ok(r#"
(module
  (func $max (export "max") (param $a i32) (param $b i32) (result i32)
    (if (result i32) (i32.gt_s (local.get $a) (local.get $b))
      (then (local.get $a))
      (else (local.get $b)))))
"#);
}
