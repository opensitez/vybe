/// End-to-end compile tests for realistic WAT programs
use super::helpers::compile_ok;

#[test]
fn fibonacci() {
    compile_ok(
        r#"
(module
  (func $fib (export "fib") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (local.get $n))
      (else
        (i32.add
          (call $fib (i32.sub (local.get $n) (i32.const 1)))
          (call $fib (i32.sub (local.get $n) (i32.const 2)))))))
)"#,
    );
}

#[test]
fn gcd() {
    compile_ok(
        r#"
(module
  (func $gcd (export "gcd") (param $a i32) (param $b i32) (result i32)
    (block $done (result i32)
      (loop $loop
        local.get $b
        i32.eqz
        br_if $done
        local.get $a
        local.get $b
        i32.rem_u
        local.get $b
        local.set $a
        local.set $b
        br $loop)
      local.get $a))
)"#,
    );
}

#[test]
fn power_of_two() {
    compile_ok(
        r#"
(module
  (func $pow2 (export "pow2") (param $n i32) (result i32)
    i32.const 1
    local.get $n
    i32.shl)
)"#,
    );
}

#[test]
fn clamp_func() {
    compile_ok(
        r#"
(module
  (func $clamp (export "clamp") (param $x i32) (param $lo i32) (param $hi i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (local.get $lo))
      (then (local.get $lo))
      (else
        (if (result i32) (i32.gt_s (local.get $x) (local.get $hi))
          (then (local.get $hi))
          (else (local.get $x))))))
)"#,
    );
}

#[test]
fn sign_func() {
    compile_ok(
        r#"
(module
  (func $sign (export "sign") (param $x i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (i32.const 0))
      (then (i32.const -1))
      (else
        (if (result i32) (i32.gt_s (local.get $x) (i32.const 0))
          (then (i32.const 1))
          (else (i32.const 0))))))
)"#,
    );
}

#[test]
fn multi_func_module() {
    compile_ok(
        r#"
(module
  (func $double (param $x i32) (result i32)
    local.get $x i32.const 2 i32.mul)
  (func $triple (param $x i32) (result i32)
    local.get $x i32.const 3 i32.mul)
  (func $sextuple (export "sextuple") (param $x i32) (result i32)
    local.get $x call $double call $triple)
)"#,
    );
}

#[test]
fn global_accumulator() {
    compile_ok(
        r#"
(module
  (global $total (mut i32) (i32.const 0))
  (func $add (export "add") (param $n i32)
    global.get $total
    local.get $n
    i32.add
    global.set $total)
  (func $get (export "get") (result i32)
    global.get $total)
  (func $reset (export "reset")
    i32.const 0
    global.set $total)
)"#,
    );
}

#[test]
fn f64_distance() {
    compile_ok(
        r#"
(module
  (func $dist (export "dist") (param $x f64) (param $y f64) (result f64)
    local.get $x local.get $x f64.mul
    local.get $y local.get $y f64.mul
    f64.add
    f64.sqrt)
)"#,
    );
}

#[test]
fn bitwise_ops() {
    compile_ok(
        r#"
(module
  (func $flags (export "flags") (param $a i32) (param $b i32) (result i32)
    local.get $a
    local.get $b
    i32.and
    local.get $a
    local.get $b
    i32.or
    i32.xor)
)"#,
    );
}

#[test]
fn type_conversions_chain() {
    compile_ok(
        r#"
(module
  (func $conv (export "conv") (param $x i32) (result f64)
    local.get $x
    f32.convert_i32_s
    f64.promote_f32)
)"#,
    );
}

#[test]
fn wast_full_script() {
    compile_ok(
        r#"
(module
  (func $add (export "add") (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.add)
  (func $sub (export "sub") (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.sub)
  (func $mul (export "mul") (param $a i32) (param $b i32) (result i32)
    local.get $a local.get $b i32.mul)
)
(assert_return (invoke "add" (i32.const 10) (i32.const 5)) (i32.const 15))
(assert_return (invoke "sub" (i32.const 10) (i32.const 5)) (i32.const 5))
(assert_return (invoke "mul" (i32.const 10) (i32.const 5)) (i32.const 50))
(assert_return (invoke "add" (i32.const 0) (i32.const 0)) (i32.const 0))
(assert_return (invoke "add" (i32.const -1) (i32.const 1)) (i32.const 0))
"#,
    );
}
