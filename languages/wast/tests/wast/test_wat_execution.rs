/// End-to-end WAT execution tests — compile + run, assert concrete output values.
/// Every test here uses `run_wast` and checks what the program actually produces.
use super::helpers::{run_wast, run_wast_one};

// ── i32 arithmetic execution ──────────────────────────────────────────────────

#[test]
fn i32_add_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 30
    i32.const 12
    i32.add
    call $log))
"#,
    );
    assert_eq!(out, "42");
}

#[test]
fn i32_mul_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 6
    i32.const 7
    i32.mul
    call $log))
"#,
    );
    assert_eq!(out, "42");
}

#[test]
fn i32_sub_negative_result() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 3
    i32.const 10
    i32.sub
    call $log))
"#,
    );
    assert_eq!(out, "-7");
}

#[test]
fn i32_div_s_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 100
    i32.const 4
    i32.div_s
    call $log))
"#,
    );
    assert_eq!(out, "25");
}

#[test]
fn i32_rem_s_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 17
    i32.const 5
    i32.rem_s
    call $log))
"#,
    );
    assert_eq!(out, "2");
}

#[test]
fn i32_bitwise_and_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 0xFF
    i32.const 0x0F
    i32.and
    call $log))
"#,
    );
    assert_eq!(out, "15");
}

#[test]
fn i32_bitwise_or_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 0xF0
    i32.const 0x0F
    i32.or
    call $log))
"#,
    );
    assert_eq!(out, "255");
}

#[test]
fn i32_shl_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 1
    i32.const 3
    i32.shl
    call $log))
"#,
    );
    assert_eq!(out, "8");
}

#[test]
fn i32_shr_u_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 128
    i32.const 2
    i32.shr_u
    call $log))
"#,
    );
    assert_eq!(out, "32");
}

// ── i32 comparison execution ──────────────────────────────────────────────────

#[test]
fn i32_eqz_true() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 0
    i32.eqz
    call $log))
"#,
    );
    assert_eq!(out, "1");
}

#[test]
fn i32_eqz_false() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 5
    i32.eqz
    call $log))
"#,
    );
    assert_eq!(out, "0");
}

#[test]
fn i32_lt_s_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 3
    i32.const 10
    i32.lt_s
    call $log))
"#,
    );
    assert_eq!(out, "1");
}

#[test]
fn i32_gt_s_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 10
    i32.const 3
    i32.gt_s
    call $log))
"#,
    );
    assert_eq!(out, "1");
}

#[test]
fn i32_eq_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 42
    i32.const 42
    i32.eq
    call $log))
"#,
    );
    assert_eq!(out, "1");
}

#[test]
fn i32_ne_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 1
    i32.const 2
    i32.ne
    call $log))
"#,
    );
    assert_eq!(out, "1");
}

// ── f64 arithmetic execution ──────────────────────────────────────────────────

#[test]
fn f64_add_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 1.5
    f64.const 2.5
    f64.add
    call $log))
"#,
    );
    assert_eq!(out, "4");
}

#[test]
fn f64_mul_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 3.0
    f64.const 4.0
    f64.mul
    call $log))
"#,
    );
    assert_eq!(out, "12");
}

#[test]
fn f64_sqrt_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 9.0
    f64.sqrt
    call $log))
"#,
    );
    assert_eq!(out, "3");
}

#[test]
fn f64_neg_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 5.0
    f64.neg
    call $log))
"#,
    );
    assert_eq!(out, "-5");
}

#[test]
fn f64_abs_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const -7.0
    f64.abs
    call $log))
"#,
    );
    assert_eq!(out, "7");
}

#[test]
fn f64_min_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 3.0
    f64.const 8.0
    f64.min
    call $log))
"#,
    );
    assert_eq!(out, "3");
}

#[test]
fn f64_max_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f64.const 3.0
    f64.const 8.0
    f64.max
    call $log))
"#,
    );
    assert_eq!(out, "8");
}

// ── locals ────────────────────────────────────────────────────────────────────

#[test]
fn local_set_get_roundtrip() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $x i32)
    i32.const 99
    local.set $x
    local.get $x
    call $log))
"#,
    );
    assert_eq!(out, "99");
}

#[test]
fn local_tee_pushes_and_sets() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $x i32)
    i32.const 55
    local.tee $x
    call $log))
"#,
    );
    assert_eq!(out, "55");
}

#[test]
fn multiple_locals_independent() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $a i32)
    (local $b i32)
    i32.const 10
    local.set $a
    i32.const 20
    local.set $b
    local.get $a
    call $log
    local.get $b
    call $log))
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

// ── globals ───────────────────────────────────────────────────────────────────

#[test]
fn global_mut_increment() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (global $g (mut i32) (i32.const 0))
  (func (export "_start")
    global.get $g
    i32.const 1
    i32.add
    global.set $g
    global.get $g
    call $log
    global.get $g
    i32.const 1
    i32.add
    global.set $g
    global.get $g
    call $log))
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn global_immutable_read() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (global $c i32 (i32.const 42))
  (func (export "_start")
    global.get $c
    call $log))
"#,
    );
    assert_eq!(out, "42");
}

// ── control flow execution ────────────────────────────────────────────────────

#[test]
fn if_then_branch_taken() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 1
    if
      i32.const 111
      call $log
    end))
"#,
    );
    assert_eq!(out, "111");
}

#[test]
fn if_else_false_branch() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 0
    if
      i32.const 1
      call $log
    else
      i32.const 2
      call $log
    end))
"#,
    );
    assert_eq!(out, "2");
}

#[test]
fn block_br_skips_rest() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    block $b
      i32.const 7
      call $log
      br $b
      i32.const 99
      call $log
    end))
"#,
    );
    assert_eq!(out, "7");
}

#[test]
fn loop_countdown() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $i i32)
    i32.const 3
    local.set $i
    block $done
      loop $again
        local.get $i
        i32.eqz
        br_if $done
        local.get $i
        call $log
        local.get $i
        i32.const 1
        i32.sub
        local.set $i
        br $again
      end
    end))
"#,
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}

#[test]
fn br_if_skips_when_false() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    block $b
      i32.const 0
      br_if $b
      i32.const 5
      call $log
    end
    i32.const 6
    call $log))
"#,
    );
    assert_eq!(out, vec!["5", "6"]);
}

// ── function calls ────────────────────────────────────────────────────────────

#[test]
fn direct_call_add() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $add (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.add)
  (func (export "_start")
    i32.const 13
    i32.const 29
    call $add
    call $log))
"#,
    );
    assert_eq!(out, "42");
}

#[test]
fn recursive_factorial() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $fact (param $n i32) (result i32)
    local.get $n
    i32.const 1
    i32.le_s
    if (result i32)
      i32.const 1
    else
      local.get $n
      local.get $n
      i32.const 1
      i32.sub
      call $fact
      i32.mul
    end)
  (func (export "_start")
    i32.const 6
    call $fact
    call $log))
"#,
    );
    assert_eq!(out, "720");
}

#[test]
fn recursive_fibonacci() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $fib (param $n i32) (result i32)
    local.get $n
    i32.const 2
    i32.lt_s
    if (result i32)
      local.get $n
    else
      local.get $n i32.const 1 i32.sub call $fib
      local.get $n i32.const 2 i32.sub call $fib
      i32.add
    end)
  (func (export "_start")
    i32.const 10
    call $fib
    call $log))
"#,
    );
    assert_eq!(out, "55");
}

#[test]
fn multi_return_values_used() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $divmod (param $a i32) (param $b i32) (result i32 i32)
    local.get $a local.get $b i32.div_u
    local.get $a local.get $b i32.rem_u)
  (func (export "_start")
    i32.const 17
    i32.const 5
    call $divmod
    call $log   ;; rem = 2
    call $log)) ;; div = 3
"#,
    );
    assert_eq!(out, vec!["2", "3"]);
}

// ── type conversions ──────────────────────────────────────────────────────────

#[test]
fn i32_to_f64_convert() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    i32.const 7
    f64.convert_i32_s
    call $log))
"#,
    );
    assert_eq!(out, "7");
}

#[test]
fn f64_to_i32_trunc() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    f64.const 3.9
    i32.trunc_f64_s
    call $log))
"#,
    );
    assert_eq!(out, "3");
}

#[test]
fn f32_to_f64_promote() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param f64)))
  (func (export "_start")
    f32.const 2.0
    f64.promote_f32
    call $log))
"#,
    );
    assert_eq!(out, "2");
}

// ── select instruction ────────────────────────────────────────────────────────

#[test]
fn select_picks_first_when_true() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 10
    i32.const 20
    i32.const 1
    select
    call $log))
"#,
    );
    assert_eq!(out, "10");
}

#[test]
fn select_picks_second_when_false() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 10
    i32.const 20
    i32.const 0
    select
    call $log))
"#,
    );
    assert_eq!(out, "20");
}

// ── br_table ──────────────────────────────────────────────────────────────────

#[test]
fn br_table_dispatch() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $dispatch (param $x i32) (result i32)
    block $default (result i32)
      block $c2 (result i32)
        block $c1 (result i32)
          block $c0 (result i32)
            local.get $x
            br_table $c0 $c1 $c2 $default
          end
          i32.const 100
          br $default
        end
        i32.const 200
        br $default
      end
      i32.const 300
      br $default
    end)
  (func (export "_start")
    i32.const 0 call $dispatch call $log
    i32.const 1 call $dispatch call $log
    i32.const 2 call $dispatch call $log))
"#,
    );
    assert_eq!(out, vec!["100", "200", "300"]);
}

// ── memory operations ─────────────────────────────────────────────────────────

#[test]
fn memory_store_load_i32() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 0     ;; address
    i32.const 12345
    i32.store
    i32.const 0
    i32.load
    call $log))
"#,
    );
    assert_eq!(out, "12345");
}

#[test]
fn memory_store_load_i8() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 4
    i32.const 200
    i32.store8
    i32.const 4
    i32.load8_u
    call $log))
"#,
    );
    assert_eq!(out, "200");
}

#[test]
fn memory_store_multiple_addresses() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (memory 1)
  (func (export "_start")
    i32.const 0  i32.const 1 i32.store
    i32.const 4  i32.const 2 i32.store
    i32.const 8  i32.const 3 i32.store
    i32.const 0 i32.load call $log
    i32.const 4 i32.load call $log
    i32.const 8 i32.load call $log))
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

// ── folded / S-expression style execution ────────────────────────────────────

#[test]
fn folded_add_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (call $log (i32.add (i32.const 19) (i32.const 23)))))
"#,
    );
    assert_eq!(out, "42");
}

#[test]
fn folded_nested_if_executed() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (call $log
      (if (result i32) (i32.const 1)
        (then (i32.const 100))
        (else (i32.const 200))))))
"#,
    );
    assert_eq!(out, "100");
}

#[test]
fn folded_local_tee_in_expr() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $x i32)
    (call $log (i32.add (local.tee $x (i32.const 10)) (i32.const 5)))))
"#,
    );
    assert_eq!(out, "15");
}

// ── nop / drop / unreachable ──────────────────────────────────────────────────

#[test]
fn nop_is_transparent() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    nop
    nop
    i32.const 5
    nop
    call $log))
"#,
    );
    assert_eq!(out, "5");
}

#[test]
fn drop_discards_value() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 999
    drop
    i32.const 1
    call $log))
"#,
    );
    assert_eq!(out, "1");
}

// ── named params / named locals ───────────────────────────────────────────────

#[test]
fn named_param_used_in_body() {
    let out = run_wast_one(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $square (param $n i32) (result i32)
    local.get $n local.get $n i32.mul)
  (func (export "_start")
    i32.const 9
    call $square
    call $log))
"#,
    );
    assert_eq!(out, "81");
}

// ── multiple outputs in one run ───────────────────────────────────────────────

#[test]
fn multiple_log_calls_in_order() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    i32.const 1 call $log
    i32.const 2 call $log
    i32.const 3 call $log
    i32.const 4 call $log
    i32.const 5 call $log))
"#,
    );
    assert_eq!(out, vec!["1", "2", "3", "4", "5"]);
}

#[test]
fn function_call_chain_values() {
    let out = run_wast(
        r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $inc (param $x i32) (result i32) local.get $x i32.const 1 i32.add)
  (func $double (param $x i32) (result i32) local.get $x i32.const 2 i32.mul)
  (func (export "_start")
    i32.const 5
    call $inc
    call $log      ;; 6
    i32.const 5
    call $double
    call $log      ;; 10
    i32.const 5
    call $inc
    call $double
    call $log))    ;; 12
"#,
    );
    assert_eq!(out, vec!["6", "10", "12"]);
}
