;; imports_exports.wat — imports, exports, inline import/export syntax
;; Run: cargo run --bin vybex -- examples/wast/imports_exports.wat

(module
  (import "wasi:cli" "log" (func $log (param i32)))

  ;; Exported constants via functions
  (func (export "pi_times_100") (result i32) (i32.const 314))
  (func (export "version")      (result i32) (i32.const 1))

  ;; Exported arithmetic
  (func (export "add") (param $a i32) (param $b i32) (result i32)
    (i32.add (local.get $a) (local.get $b)))

  (func (export "mul") (param $a i32) (param $b i32) (result i32)
    (i32.mul (local.get $a) (local.get $b)))

  ;; Exported mutable global
  (global $g (export "counter") (mut i32) (i32.const 0))

  (func (export "inc") (global.set $g (i32.add (global.get $g) (i32.const 1))))
  (func (export "get") (result i32) (global.get $g))

  (func $demo (export "demo")
    (call $log (i32.const 314))   ;; pi_times_100
    (call $log (i32.const 1))     ;; version
    (call $log (i32.add (i32.const 10) (i32.const 5)))  ;; 15
    (call $log (i32.mul (i32.const 6)  (i32.const 7)))  ;; 42
    (global.set $g (i32.const 0))
    (global.set $g (i32.add (global.get $g) (i32.const 1)))
    (global.set $g (i32.add (global.get $g) (i32.const 1)))
    (call $log (global.get $g)))  ;; 2

  (start $demo)
)
