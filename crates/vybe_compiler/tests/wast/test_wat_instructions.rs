/// Tests for WAT instructions — all numeric ops, control flow, memory, variables
use super::helpers::{parse_ok, compile_ok};

// ── i32 arithmetic ────────────────────────────────────────────────────────────

#[test] fn i32_const()   { parse_ok("(module (func (result i32) i32.const 42))"); }
#[test] fn i32_add()     { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))"); }
#[test] fn i32_sub()     { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.sub))"); }
#[test] fn i32_mul()     { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.mul))"); }
#[test] fn i32_div_s()   { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.div_s))"); }
#[test] fn i32_div_u()   { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.div_u))"); }
#[test] fn i32_rem_s()   { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.rem_s))"); }
#[test] fn i32_rem_u()   { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.rem_u))"); }
#[test] fn i32_and()     { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.and))"); }
#[test] fn i32_or()      { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.or))"); }
#[test] fn i32_xor()     { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.xor))"); }
#[test] fn i32_shl()     { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.shl))"); }
#[test] fn i32_shr_s()   { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.shr_s))"); }
#[test] fn i32_shr_u()   { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.shr_u))"); }
#[test] fn i32_rotl()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.rotl))"); }
#[test] fn i32_rotr()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.rotr))"); }
#[test] fn i32_clz()     { parse_ok("(module (func (param i32) (result i32) local.get 0 i32.clz))"); }
#[test] fn i32_ctz()     { parse_ok("(module (func (param i32) (result i32) local.get 0 i32.ctz))"); }
#[test] fn i32_popcnt()  { parse_ok("(module (func (param i32) (result i32) local.get 0 i32.popcnt))"); }

// ── i32 comparisons ───────────────────────────────────────────────────────────

#[test] fn i32_eqz()     { parse_ok("(module (func (param i32) (result i32) local.get 0 i32.eqz))"); }
#[test] fn i32_eq()      { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.eq))"); }
#[test] fn i32_ne()      { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.ne))"); }
#[test] fn i32_lt_s()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.lt_s))"); }
#[test] fn i32_lt_u()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.lt_u))"); }
#[test] fn i32_gt_s()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.gt_s))"); }
#[test] fn i32_gt_u()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.gt_u))"); }
#[test] fn i32_le_s()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.le_s))"); }
#[test] fn i32_le_u()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.le_u))"); }
#[test] fn i32_ge_s()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.ge_s))"); }
#[test] fn i32_ge_u()    { parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.ge_u))"); }

// ── i64 arithmetic ────────────────────────────────────────────────────────────

#[test] fn i64_const()   { parse_ok("(module (func (result i64) i64.const 9999999999))"); }
#[test] fn i64_add()     { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.add))"); }
#[test] fn i64_sub()     { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.sub))"); }
#[test] fn i64_mul()     { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.mul))"); }
#[test] fn i64_div_s()   { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.div_s))"); }
#[test] fn i64_rem_s()   { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.rem_s))"); }
#[test] fn i64_and()     { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.and))"); }
#[test] fn i64_or()      { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.or))"); }
#[test] fn i64_xor()     { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.xor))"); }
#[test] fn i64_shl()     { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.shl))"); }
#[test] fn i64_shr_s()   { parse_ok("(module (func (param i64 i64) (result i64) local.get 0 local.get 1 i64.shr_s))"); }
#[test] fn i64_clz()     { parse_ok("(module (func (param i64) (result i64) local.get 0 i64.clz))"); }
#[test] fn i64_eqz()     { parse_ok("(module (func (param i64) (result i32) local.get 0 i64.eqz))"); }
#[test] fn i64_eq()      { parse_ok("(module (func (param i64 i64) (result i32) local.get 0 local.get 1 i64.eq))"); }
#[test] fn i64_lt_s()    { parse_ok("(module (func (param i64 i64) (result i32) local.get 0 local.get 1 i64.lt_s))"); }
#[test] fn i64_ge_u()    { parse_ok("(module (func (param i64 i64) (result i32) local.get 0 local.get 1 i64.ge_u))"); }

// ── f32 arithmetic ────────────────────────────────────────────────────────────

#[test] fn f32_const()   { parse_ok("(module (func (result f32) f32.const 3.14))"); }
#[test] fn f32_add()     { parse_ok("(module (func (param f32 f32) (result f32) local.get 0 local.get 1 f32.add))"); }
#[test] fn f32_sub()     { parse_ok("(module (func (param f32 f32) (result f32) local.get 0 local.get 1 f32.sub))"); }
#[test] fn f32_mul()     { parse_ok("(module (func (param f32 f32) (result f32) local.get 0 local.get 1 f32.mul))"); }
#[test] fn f32_div()     { parse_ok("(module (func (param f32 f32) (result f32) local.get 0 local.get 1 f32.div))"); }
#[test] fn f32_abs()     { parse_ok("(module (func (param f32) (result f32) local.get 0 f32.abs))"); }
#[test] fn f32_neg()     { parse_ok("(module (func (param f32) (result f32) local.get 0 f32.neg))"); }
#[test] fn f32_sqrt()    { parse_ok("(module (func (param f32) (result f32) local.get 0 f32.sqrt))"); }
#[test] fn f32_ceil()    { parse_ok("(module (func (param f32) (result f32) local.get 0 f32.ceil))"); }
#[test] fn f32_floor()   { parse_ok("(module (func (param f32) (result f32) local.get 0 f32.floor))"); }
#[test] fn f32_trunc()   { parse_ok("(module (func (param f32) (result f32) local.get 0 f32.trunc))"); }
#[test] fn f32_nearest() { parse_ok("(module (func (param f32) (result f32) local.get 0 f32.nearest))"); }
#[test] fn f32_min()     { parse_ok("(module (func (param f32 f32) (result f32) local.get 0 local.get 1 f32.min))"); }
#[test] fn f32_max()     { parse_ok("(module (func (param f32 f32) (result f32) local.get 0 local.get 1 f32.max))"); }
#[test] fn f32_copysign(){ parse_ok("(module (func (param f32 f32) (result f32) local.get 0 local.get 1 f32.copysign))"); }
#[test] fn f32_eq()      { parse_ok("(module (func (param f32 f32) (result i32) local.get 0 local.get 1 f32.eq))"); }
#[test] fn f32_ne()      { parse_ok("(module (func (param f32 f32) (result i32) local.get 0 local.get 1 f32.ne))"); }
#[test] fn f32_lt()      { parse_ok("(module (func (param f32 f32) (result i32) local.get 0 local.get 1 f32.lt))"); }
#[test] fn f32_gt()      { parse_ok("(module (func (param f32 f32) (result i32) local.get 0 local.get 1 f32.gt))"); }
#[test] fn f32_le()      { parse_ok("(module (func (param f32 f32) (result i32) local.get 0 local.get 1 f32.le))"); }
#[test] fn f32_ge()      { parse_ok("(module (func (param f32 f32) (result i32) local.get 0 local.get 1 f32.ge))"); }

// ── f64 arithmetic ────────────────────────────────────────────────────────────

#[test] fn f64_const()   { parse_ok("(module (func (result f64) f64.const 2.718281828))"); }
#[test] fn f64_add()     { parse_ok("(module (func (param f64 f64) (result f64) local.get 0 local.get 1 f64.add))"); }
#[test] fn f64_mul()     { parse_ok("(module (func (param f64 f64) (result f64) local.get 0 local.get 1 f64.mul))"); }
#[test] fn f64_div()     { parse_ok("(module (func (param f64 f64) (result f64) local.get 0 local.get 1 f64.div))"); }
#[test] fn f64_abs()     { parse_ok("(module (func (param f64) (result f64) local.get 0 f64.abs))"); }
#[test] fn f64_neg()     { parse_ok("(module (func (param f64) (result f64) local.get 0 f64.neg))"); }
#[test] fn f64_sqrt()    { parse_ok("(module (func (param f64) (result f64) local.get 0 f64.sqrt))"); }
#[test] fn f64_ceil()    { parse_ok("(module (func (param f64) (result f64) local.get 0 f64.ceil))"); }
#[test] fn f64_floor()   { parse_ok("(module (func (param f64) (result f64) local.get 0 f64.floor))"); }
#[test] fn f64_nearest() { parse_ok("(module (func (param f64) (result f64) local.get 0 f64.nearest))"); }
#[test] fn f64_min()     { parse_ok("(module (func (param f64 f64) (result f64) local.get 0 local.get 1 f64.min))"); }
#[test] fn f64_max()     { parse_ok("(module (func (param f64 f64) (result f64) local.get 0 local.get 1 f64.max))"); }
#[test] fn f64_eq()      { parse_ok("(module (func (param f64 f64) (result i32) local.get 0 local.get 1 f64.eq))"); }
#[test] fn f64_lt()      { parse_ok("(module (func (param f64 f64) (result i32) local.get 0 local.get 1 f64.lt))"); }
#[test] fn f64_ge()      { parse_ok("(module (func (param f64 f64) (result i32) local.get 0 local.get 1 f64.ge))"); }

// ── Special float literals ────────────────────────────────────────────────────

#[test] fn f32_const_inf()      { parse_ok("(module (func (result f32) f32.const inf))"); }
#[test] fn f32_const_neg_inf()  { parse_ok("(module (func (result f32) f32.const -inf))"); }
#[test] fn f32_const_nan()      { parse_ok("(module (func (result f32) f32.const nan))"); }
#[test] fn f32_const_nan_hex()  { parse_ok("(module (func (result f32) f32.const nan:0x200000))"); }
#[test] fn f64_const_inf()      { parse_ok("(module (func (result f64) f64.const inf))"); }
#[test] fn f64_const_neg_inf()  { parse_ok("(module (func (result f64) f64.const -inf))"); }
#[test] fn f64_const_nan()      { parse_ok("(module (func (result f64) f64.const nan))"); }
#[test] fn f64_const_hex_float(){ parse_ok("(module (func (result f64) f64.const 0x1.8p1))"); }

// ── Hex integer literals ──────────────────────────────────────────────────────

#[test] fn i32_const_hex()      { parse_ok("(module (func (result i32) i32.const 0xff))"); }
#[test] fn i32_const_neg_hex()  { parse_ok("(module (func (result i32) i32.const -0x80000000))"); }
#[test] fn i64_const_hex()      { parse_ok("(module (func (result i64) i64.const 0xffffffffffffffff))"); }

// ── Type conversions ──────────────────────────────────────────────────────────

#[test] fn i32_wrap_i64()         { parse_ok("(module (func (param i64) (result i32) local.get 0 i32.wrap_i64))"); }
#[test] fn i64_extend_i32_s()     { parse_ok("(module (func (param i32) (result i64) local.get 0 i64.extend_i32_s))"); }
#[test] fn i64_extend_i32_u()     { parse_ok("(module (func (param i32) (result i64) local.get 0 i64.extend_i32_u))"); }
#[test] fn f32_convert_i32_s()    { parse_ok("(module (func (param i32) (result f32) local.get 0 f32.convert_i32_s))"); }
#[test] fn f64_convert_i32_u()    { parse_ok("(module (func (param i32) (result f64) local.get 0 f64.convert_i32_u))"); }
#[test] fn f32_demote_f64()       { parse_ok("(module (func (param f64) (result f32) local.get 0 f32.demote_f64))"); }
#[test] fn f64_promote_f32()      { parse_ok("(module (func (param f32) (result f64) local.get 0 f64.promote_f32))"); }
#[test] fn i32_trunc_f32_s()      { parse_ok("(module (func (param f32) (result i32) local.get 0 i32.trunc_f32_s))"); }
#[test] fn i32_trunc_sat_f64_u()  { parse_ok("(module (func (param f64) (result i32) local.get 0 i32.trunc_sat_f64_u))"); }
#[test] fn i32_reinterpret_f32()  { parse_ok("(module (func (param f32) (result i32) local.get 0 i32.reinterpret_f32))"); }
#[test] fn f64_reinterpret_i64()  { parse_ok("(module (func (param i64) (result f64) local.get 0 f64.reinterpret_i64))"); }
#[test] fn i32_extend8_s()        { parse_ok("(module (func (param i32) (result i32) local.get 0 i32.extend8_s))"); }
#[test] fn i64_extend32_s()       { parse_ok("(module (func (param i64) (result i64) local.get 0 i64.extend32_s))"); }

// ── Variable instructions ─────────────────────────────────────────────────────

#[test] fn local_get_by_index()  { parse_ok("(module (func (param i32) (result i32) local.get 0))"); }
#[test] fn local_get_by_name()   { parse_ok("(module (func (param $x i32) (result i32) local.get $x))"); }
#[test] fn local_set()           { parse_ok("(module (func (param i32) (local i32) local.get 0 local.set 1))"); }
#[test] fn local_tee()           { parse_ok("(module (func (param i32) (result i32) (local i32) local.get 0 local.tee 1))"); }
#[test] fn global_get()          { parse_ok("(module (global $g i32 (i32.const 0)) (func (result i32) global.get $g))"); }
#[test] fn global_set()          { parse_ok("(module (global $g (mut i32) (i32.const 0)) (func i32.const 1 global.set $g))"); }

// ── Parametric instructions ───────────────────────────────────────────────────

#[test] fn drop_instr()   { parse_ok("(module (func i32.const 1 drop))"); }
#[test] fn select_instr() { parse_ok("(module (func (param i32 i32 i32) (result i32) local.get 0 local.get 1 local.get 2 select))"); }
#[test] fn nop_instr()    { parse_ok("(module (func nop))"); }
#[test] fn unreachable()  { parse_ok("(module (func unreachable))"); }

// ── Control flow — plain form ─────────────────────────────────────────────────

#[test]
fn return_instr() {
    parse_ok("(module (func (result i32) i32.const 42 return))");
}

#[test]
fn block_plain() {
    parse_ok("(module (func block nop end))");
}

#[test]
fn block_with_result() {
    parse_ok("(module (func (result i32) block (result i32) i32.const 1 end))");
}

#[test]
fn loop_plain() {
    parse_ok("(module (func loop nop end))");
}

#[test]
fn if_then_else_plain() {
    parse_ok("(module (func (param i32) (result i32) local.get 0 if (result i32) i32.const 1 else i32.const 0 end))");
}

#[test]
fn br_instr() {
    parse_ok("(module (func block br 0 end))");
}

#[test]
fn br_if_instr() {
    parse_ok("(module (func (param i32) block local.get 0 br_if 0 end))");
}

#[test]
fn br_table_instr() {
    parse_ok("(module (func (param i32) block block local.get 0 br_table 0 1 end end))");
}

#[test]
fn call_instr() {
    parse_ok("(module (func $f (param i32) (result i32) local.get 0) (func (result i32) i32.const 5 call $f))");
}

#[test]
fn call_indirect_instr() {
    parse_ok("(module (type $t (func (param i32) (result i32))) (table 1 funcref) (func (param i32 i32) (result i32) local.get 0 local.get 1 call_indirect (type $t)))");
}

// ── Control flow — folded form ────────────────────────────────────────────────

#[test]
fn folded_if_then() {
    parse_ok("(module (func (param i32) (if (local.get 0) (then nop))))");
}

#[test]
fn folded_if_then_else() {
    parse_ok("(module (func (param i32) (result i32) (if (result i32) (local.get 0) (then (i32.const 1)) (else (i32.const 0)))))");
}

#[test]
fn folded_block() {
    parse_ok("(module (func (block $b (br $b))))");
}

#[test]
fn folded_loop() {
    parse_ok("(module (func (loop $l (br $l))))");
}

#[test]
fn folded_add() {
    parse_ok("(module (func (param i32 i32) (result i32) (i32.add (local.get 0) (local.get 1))))");
}

#[test]
fn deeply_nested_folded() {
    parse_ok("(module (func (param i32 i32 i32) (result i32) (i32.add (i32.mul (local.get 0) (local.get 1)) (local.get 2))))");
}

// ── Memory instructions ───────────────────────────────────────────────────────

#[test]
fn memory_size_instr() {
    parse_ok("(module (memory 1) (func (result i32) memory.size))");
}

#[test]
fn memory_grow_instr() {
    parse_ok("(module (memory 1) (func (param i32) (result i32) local.get 0 memory.grow))");
}

#[test]
fn i32_load_instr() {
    parse_ok("(module (memory 1) (func (param i32) (result i32) local.get 0 i32.load))");
}

#[test]
fn i32_load_with_offset() {
    parse_ok("(module (memory 1) (func (param i32) (result i32) local.get 0 i32.load offset=4))");
}

#[test]
fn i32_load_with_align() {
    parse_ok("(module (memory 1) (func (param i32) (result i32) local.get 0 i32.load align=4))");
}

#[test]
fn i32_store_instr() {
    parse_ok("(module (memory 1) (func (param i32 i32) local.get 0 local.get 1 i32.store))");
}

#[test]
fn i64_load32_s() {
    parse_ok("(module (memory 1) (func (param i32) (result i64) local.get 0 i64.load32_s))");
}

#[test]
fn i32_load8_u() {
    parse_ok("(module (memory 1) (func (param i32) (result i32) local.get 0 i32.load8_u))");
}

#[test]
fn f64_store_instr() {
    parse_ok("(module (memory 1) (func (param i32 f64) local.get 0 local.get 1 f64.store))");
}

// ── Reference types ───────────────────────────────────────────────────────────

#[test]
fn ref_null_funcref() {
    parse_ok("(module (func ref.null funcref))");
}

#[test]
fn ref_is_null() {
    parse_ok("(module (func (param funcref) (result i32) local.get 0 ref.is_null))");
}

#[test]
fn ref_func_instr() {
    parse_ok("(module (func $f) (func ref.func $f drop))");
}

// ── Comments ──────────────────────────────────────────────────────────────────

#[test]
fn line_comment() {
    parse_ok("(module ;; this is a comment\n (func))");
}

#[test]
fn block_comment() {
    parse_ok("(module (; block comment ;) (func))");
}

#[test]
fn nested_block_comment() {
    parse_ok("(module (; outer (; inner ;) outer ;) (func))");
}

// ── Compile checks ────────────────────────────────────────────────────────────

#[test]
fn compile_add_func() {
    compile_ok("(module (func $add (export \"add\") (param $a i32) (param $b i32) (result i32) local.get $a local.get $b i32.add))");
}

#[test]
fn compile_factorial_folded() {
    compile_ok(r#"
(module
  (func $factorial (export "factorial") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (i32.const 1))
      (else (i32.mul (local.get $n)
                     (call $factorial (i32.sub (local.get $n) (i32.const 1)))))))
)"#);
}

#[test]
fn compile_global_counter() {
    compile_ok(r#"
(module
  (global $count (mut i32) (i32.const 0))
  (func $inc (export "inc")
    global.get $count
    i32.const 1
    i32.add
    global.set $count)
  (func $get (export "get") (result i32)
    global.get $count)
)"#);
}
