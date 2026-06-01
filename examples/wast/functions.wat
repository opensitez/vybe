;; functions.wat — function calls, recursion, locals, select
;; Run: cargo run --bin vybex -- examples/wast/functions.wat

(module
  (import "wasi:cli" "log"  (func $log  (param i32)))
  (import "wasi:cli" "logf" (func $logf (param f64)))

  ;; Fibonacci (recursive)
  (func $fib (export "fib") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (local.get $n))
      (else
        (i32.add
          (call $fib (i32.sub (local.get $n) (i32.const 1)))
          (call $fib (i32.sub (local.get $n) (i32.const 2)))))))

  ;; Mutual recursion
  (func $is_even (export "is_even") (param $n i32) (result i32)
    (if (result i32) (i32.eqz (local.get $n))
      (then (i32.const 1))
      (else (call $is_odd (i32.sub (local.get $n) (i32.const 1))))))

  (func $is_odd (export "is_odd") (param $n i32) (result i32)
    (if (result i32) (i32.eqz (local.get $n))
      (then (i32.const 0))
      (else (call $is_even (i32.sub (local.get $n) (i32.const 1))))))

  ;; local.tee: x*x + x
  (func $poly (export "poly") (param $x i32) (result i32)
    (local $t i32)
    (i32.add
      (i32.mul (local.get $x) (local.tee $t (local.get $x)))
      (local.get $t)))

  ;; (a+b) * (a-b) = a^2 - b^2
  (func $diff_sq (export "diff_sq") (param $a i32) (param $b i32) (result i32)
    (i32.mul
      (i32.add (local.get $a) (local.get $b))
      (i32.sub (local.get $a) (local.get $b))))

  ;; Branchless max via select
  (func $max (export "max") (param $a i32) (param $b i32) (result i32)
    (select
      (local.get $a)
      (local.get $b)
      (i32.gt_s (local.get $a) (local.get $b))))

  ;; Float hypotenuse: sqrt(a^2 + b^2)
  (func $hypot (export "hypot") (param $a f64) (param $b f64) (result f64)
    (f64.sqrt
      (f64.add
        (f64.mul (local.get $a) (local.get $a))
        (f64.mul (local.get $b) (local.get $b)))))

  (func $demo (export "demo")
    (call $log  (call $fib     (i32.const 10)))   ;; 55
    (call $log  (call $is_even (i32.const 4)))    ;; 1
    (call $log  (call $is_odd  (i32.const 7)))    ;; 1
    (call $log  (call $poly    (i32.const 5)))    ;; 30
    (call $log  (call $diff_sq (i32.const 7) (i32.const 3))) ;; 40
    (call $log  (call $max     (i32.const 3) (i32.const 9))) ;; 9
    (call $logf (call $hypot   (f64.const 3.0) (f64.const 4.0)))) ;; 5.0

  (start $demo)
)
