;; control_flow.wat — structured control flow: if/else, recursion, select
;; Run: cargo run --bin vybex -- examples/wast/control_flow.wat

(module
  (import "wasi:cli" "log" (func $log (param i32)))

  ;; Factorial — recursive
  (func $factorial (export "factorial") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 1))
      (then (i32.const 1))
      (else (i32.mul (local.get $n)
              (call $factorial (i32.sub (local.get $n) (i32.const 1)))))))

  ;; Triangular number: 1+2+...+n — recursive (avoid name "sum")
  (func $tri (export "tri") (param $n i32) (result i32)
    (if (result i32) (i32.le_s (local.get $n) (i32.const 0))
      (then (i32.const 0))
      (else (i32.add (local.get $n)
              (call $tri (i32.sub (local.get $n) (i32.const 1)))))))

  ;; GCD — recursive Euclidean
  (func $gcd (export "gcd") (param $a i32) (param $b i32) (result i32)
    (if (result i32) (i32.eqz (local.get $b))
      (then (local.get $a))
      (else (call $gcd (local.get $b) (i32.rem_u (local.get $a) (local.get $b))))))

  ;; Clamp — nested if/else
  (func $clamp (export "clamp") (param $x i32) (param $lo i32) (param $hi i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (local.get $lo))
      (then (local.get $lo))
      (else
        (if (result i32) (i32.gt_s (local.get $x) (local.get $hi))
          (then (local.get $hi))
          (else (local.get $x))))))

  ;; Branchless max via select
  (func $max (export "max") (param $a i32) (param $b i32) (result i32)
    (select
      (local.get $a)
      (local.get $b)
      (i32.gt_s (local.get $a) (local.get $b))))

  (func $demo (export "demo")
    (call $log (call $factorial (i32.const 5)))    ;; 120
    (call $log (call $factorial (i32.const 10)))   ;; 3628800
    (call $log (call $tri       (i32.const 10)))   ;; 55
    (call $log (call $tri       (i32.const 100)))  ;; 5050
    (call $log (call $gcd       (i32.const 48) (i32.const 18))) ;; 6
    (call $log (call $clamp     (i32.const 15) (i32.const 0) (i32.const 10))) ;; 10
    (call $log (call $max       (i32.const 3)  (i32.const 9)))) ;; 9

  (start $demo)
)
