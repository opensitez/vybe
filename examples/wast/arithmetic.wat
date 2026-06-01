;; arithmetic.wat — basic integer and float arithmetic
;; Run: cargo run --bin vybex -- examples/wast/arithmetic.wat

(module
  (import "wasi:cli" "log"  (func $log  (param i32)))
  (import "wasi:cli" "logf" (func $logf (param f64)))

  (func $add (export "add") (param $a i32) (param $b i32) (result i32)
    (i32.add (local.get $a) (local.get $b)))

  (func $sub (export "sub") (param $a i32) (param $b i32) (result i32)
    (i32.sub (local.get $a) (local.get $b)))

  (func $mul (export "mul") (param $a i32) (param $b i32) (result i32)
    (i32.mul (local.get $a) (local.get $b)))

  (func $div (export "div") (param $a i32) (param $b i32) (result i32)
    (i32.div_s (local.get $a) (local.get $b)))

  (func $rem (export "rem") (param $a i32) (param $b i32) (result i32)
    (i32.rem_s (local.get $a) (local.get $b)))

  (func $abs (export "abs") (param $x i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (i32.const 0))
      (then (i32.sub (i32.const 0) (local.get $x)))
      (else (local.get $x))))

  (func $max (export "max") (param $a i32) (param $b i32) (result i32)
    (if (result i32) (i32.gt_s (local.get $a) (local.get $b))
      (then (local.get $a))
      (else (local.get $b))))

  (func $min (export "min") (param $a i32) (param $b i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $a) (local.get $b))
      (then (local.get $a))
      (else (local.get $b))))

  (func $fadd (export "fadd") (param $a f64) (param $b f64) (result f64)
    (f64.add (local.get $a) (local.get $b)))

  (func $fsqrt (export "fsqrt") (param $x f64) (result f64)
    (f64.sqrt (local.get $x)))

  (func $pow2 (export "pow2") (param $n i32) (result i32)
    (i32.shl (i32.const 1) (local.get $n)))

  (func $clamp (export "clamp") (param $x i32) (param $lo i32) (param $hi i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (local.get $lo))
      (then (local.get $lo))
      (else
        (if (result i32) (i32.gt_s (local.get $x) (local.get $hi))
          (then (local.get $hi))
          (else (local.get $x))))))

  (func $demo (export "demo")
    (call $log  (call $add   (i32.const 10)  (i32.const 3)))   ;; 13
    (call $log  (call $sub   (i32.const 10)  (i32.const 3)))   ;; 7
    (call $log  (call $mul   (i32.const 6)   (i32.const 7)))   ;; 42
    (call $log  (call $div   (i32.const 20)  (i32.const 4)))   ;; 5
    (call $log  (call $rem   (i32.const 17)  (i32.const 5)))   ;; 2
    (call $log  (call $abs   (i32.const -42)))                 ;; 42
    (call $log  (call $max   (i32.const 3)   (i32.const 7)))   ;; 7
    (call $log  (call $min   (i32.const 3)   (i32.const 7)))   ;; 3
    (call $logf (call $fadd  (f64.const 1.5) (f64.const 2.5))) ;; 4
    (call $logf (call $fsqrt (f64.const 2.0)))                 ;; 1.4142135623730951
    (call $log  (call $pow2  (i32.const 8)))                   ;; 256
    (call $log  (call $clamp (i32.const 15)  (i32.const 0) (i32.const 10)))) ;; 10

  (start $demo)
)
