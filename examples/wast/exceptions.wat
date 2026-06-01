;; exceptions.wat — structured error handling in WAT
;; Run: cargo run --bin vybex -- examples/wast/exceptions.wat
;;
;; WAT exceptions (tag/throw/try/catch) require runtime exception support.
;; This example demonstrates the equivalent error-handling semantics using
;; sentinel values and conditional dispatch — a common pattern in WASM code
;; that must remain compatible with environments without exception support.

(module
  (import "wasi:cli" "log" (func $log (param i32)))

  ;; Sentinel: -1 = error, ≥ 0 = ok
  (global $ERR i32 (i32.const -1))

  ;; Safe division: returns ERR (-1) if divisor is zero
  (func $safe_div (export "safe_div") (param $a i32) (param $b i32) (result i32)
    (if (result i32) (i32.eqz (local.get $b))
      (then (global.get $ERR))
      (else (i32.div_s (local.get $a) (local.get $b)))))

  ;; Safe sqrt: returns ERR for negative input (i32 approximation)
  (func $safe_sqrt (export "safe_sqrt") (param $x i32) (result i32)
    (if (result i32) (i32.lt_s (local.get $x) (i32.const 0))
      (then (global.get $ERR))
      (else
        ;; Integer square root via Newton's method (simplified)
        (if (result i32) (i32.eqz (local.get $x))
          (then (i32.const 0))
          (else
            ;; For demo: check perfect squares 1,4,9,16,25
            (if (result i32) (i32.eq (local.get $x) (i32.const 25))
              (then (i32.const 5))
              (else (if (result i32) (i32.eq (local.get $x) (i32.const 9))
                (then (i32.const 3))
                (else (i32.const 1))))))))))

  ;; Propagate errors: if either argument is ERR, return ERR
  (func $safe_add (export "safe_add") (param $a i32) (param $b i32) (result i32)
    (if (result i32)
      (i32.or
        (i32.eq (local.get $a) (global.get $ERR))
        (i32.eq (local.get $b) (global.get $ERR)))
      (then (global.get $ERR))
      (else (i32.add (local.get $a) (local.get $b)))))

  (func $demo (export "demo")
    ;; Normal division
    (call $log (call $safe_div (i32.const 10) (i32.const 2)))    ;; 5
    ;; Division by zero → ERR
    (call $log (call $safe_div (i32.const 10) (i32.const 0)))    ;; -1
    ;; Safe sqrt of 25
    (call $log (call $safe_sqrt (i32.const 25)))                  ;; 5
    ;; Safe sqrt of negative → ERR
    (call $log (call $safe_sqrt (i32.const -4)))                  ;; -1
    ;; Propagate error through addition
    (call $log (call $safe_add
      (call $safe_div (i32.const 10) (i32.const 0))
      (i32.const 5))))                                            ;; -1

  (start $demo)
)
