;; tables_indirect.wat — funcref tables and indirect function dispatch
;; Run: cargo run --bin vybex -- examples/wast/tables_indirect.wat
;;
;; WAT tables store function references for runtime dispatch. This example
;; uses a switch pattern to demonstrate the dispatch semantics; full
;; call_indirect (table-based indirect calls) requires runtime table support.

(module
  (import "wasi:cli" "log" (func $log (param i32)))

  (type $unary (func (param i32) (result i32)))

  ;; Functions to dispatch to
  (func $double (export "double") (param $x i32) (result i32)
    (i32.mul (local.get $x) (i32.const 2)))

  (func $triple (export "triple") (param $x i32) (result i32)
    (i32.mul (local.get $x) (i32.const 3)))

  (func $square (export "square") (param $x i32) (result i32)
    (i32.mul (local.get $x) (local.get $x)))

  (func $negate (export "negate") (param $x i32) (result i32)
    (i32.sub (i32.const 0) (local.get $x)))

  ;; Software dispatch table: select function by index
  (func $dispatch (export "dispatch") (param $fn_idx i32) (param $x i32) (result i32)
    (if (result i32) (i32.eq (local.get $fn_idx) (i32.const 0))
      (then (call $double (local.get $x)))
      (else (if (result i32) (i32.eq (local.get $fn_idx) (i32.const 1))
        (then (call $triple (local.get $x)))
        (else (if (result i32) (i32.eq (local.get $fn_idx) (i32.const 2))
          (then (call $square (local.get $x)))
          (else (call $negate (local.get $x)))))))))

  ;; Apply a function over a range [lo..hi], accumulate results
  (func $map_reduce (export "map_reduce")
        (param $fn_idx i32) (param $lo i32) (param $hi i32) (result i32)
    (if (result i32) (i32.gt_s (local.get $lo) (local.get $hi))
      (then (i32.const 0))
      (else
        (i32.add
          (call $dispatch (local.get $fn_idx) (local.get $lo))
          (call $map_reduce
            (local.get $fn_idx)
            (i32.add (local.get $lo) (i32.const 1))
            (local.get $hi))))))

  (func $demo (export "demo")
    ;; Direct dispatch
    (call $log (call $dispatch (i32.const 0) (i32.const 5)))   ;; double(5)  = 10
    (call $log (call $dispatch (i32.const 1) (i32.const 5)))   ;; triple(5)  = 15
    (call $log (call $dispatch (i32.const 2) (i32.const 5)))   ;; square(5)  = 25
    (call $log (call $dispatch (i32.const 3) (i32.const 5)))   ;; negate(5)  = -5
    ;; map double over [1..5]: 2+4+6+8+10 = 30
    (call $log (call $map_reduce (i32.const 0) (i32.const 1) (i32.const 5)))) ;; 30

  (start $demo)
)
