;; simd.wat — SIMD v128 via splat and vector arithmetic
;; Run: cargo run --bin vybex -- examples/wast/simd.wat

(module
  (import "wasi:cli" "log" (func $log (param i32)))

  ;; Multiply all 4 i32 lanes by 2, check sum of lanes via repeated splat+add
  ;; We verify SIMD correctness by checking any_true after subtraction
  ;;
  ;; Pattern: splat(a)+splat(b) == splat(a+b)  →  sub == 0  →  any_true=0
  (func $add_correct (export "add_correct") (param $a i32) (param $b i32) (result i32)
    ;; (splat(a) + splat(b)) - splat(a+b) should be zero vector → any_true=0
    ;; We return 1 if add is correct (any_true of zero = 0, so negate)
    (i32.sub (i32.const 1)
      (v128.any_true
        (i32x4.sub
          (i32x4.add (i32x4.splat (local.get $a)) (i32x4.splat (local.get $b)))
          (i32x4.splat (i32.add (local.get $a) (local.get $b)))))))

  ;; f32x4 sqrt: splat(4.0), take sqrt of each lane, check > 0 via any_true
  (func $f32_sqrt_positive (export "f32_sqrt_positive") (result i32)
    (v128.any_true
      (f32x4.sqrt (f32x4.splat (f32.const 4.0)))))

  ;; v128 bitwise AND: any_true of (splat(0xF0) AND splat(0x0F)) = 0 (no overlap)
  (func $and_disjoint (export "and_disjoint") (result i32)
    (v128.any_true
      (v128.and
        (i32x4.splat (i32.const 0xF0))
        (i32x4.splat (i32.const 0x0F)))))

  ;; v128 OR: any_true of (splat(0) OR splat(1)) = 1
  (func $or_nonzero (export "or_nonzero") (result i32)
    (v128.any_true
      (v128.or
        (i32x4.splat (i32.const 0))
        (i32x4.splat (i32.const 1)))))

  (func $demo (export "demo")
    (call $log (call $add_correct        (i32.const 7) (i32.const 5))) ;; 1
    (call $log (call $f32_sqrt_positive))  ;; 1
    (call $log (call $and_disjoint))       ;; 0
    (call $log (call $or_nonzero)))        ;; 1

  (start $demo)
)
