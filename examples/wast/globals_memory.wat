;; globals_memory.wat — mutable globals
;; Run: cargo run --bin vybex -- examples/wast/globals_memory.wat

(module
  (import "wasi:cli" "log" (func $log (param i32)))

  ;; Mutable global counter — persists across function calls
  (global $counter (mut i32) (i32.const 0))
  (global $total   (mut i32) (i32.const 0))

  (func $inc (export "inc")
    (global.set $counter (i32.add (global.get $counter) (i32.const 1))))

  (func $add_to_total (export "add_to_total") (param $n i32)
    (global.set $total (i32.add (global.get $total) (local.get $n))))

  (func $reset (export "reset")
    (global.set $counter (i32.const 0))
    (global.set $total   (i32.const 0)))

  (func $get_counter (export "get_counter") (result i32)
    (global.get $counter))

  (func $get_total (export "get_total") (result i32)
    (global.get $total))

  (func $demo (export "demo")
    ;; Counter starts at 0
    (call $log (call $get_counter))    ;; 0
    ;; Increment 5 times
    (call $inc) (call $inc) (call $inc) (call $inc) (call $inc)
    (call $log (call $get_counter))    ;; 5
    ;; Accumulate into total
    (call $add_to_total (i32.const 10))
    (call $add_to_total (i32.const 20))
    (call $add_to_total (i32.const 12))
    (call $log (call $get_total))      ;; 42
    ;; Reset both
    (call $reset)
    (call $log (call $get_counter))    ;; 0
    (call $log (call $get_total)))     ;; 0

  (start $demo)
)
