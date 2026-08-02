;; vybe-test: wast/wat_loops/test_loop_with_continue_pattern
;; origin: languages/wast/tests/wast/test_wat_loops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $i i32) (local $s i32)
        block loop
          local.get $i i32.const 5 i32.ge_s br_if 1
          local.get $i i32.const 1 i32.add local.set $i
          local.get $i i32.const 3 i32.eq
          if br 1 end
          local.get $s local.get $i i32.add local.set $s
          br 0
        end end local.get $s call $log)
)
