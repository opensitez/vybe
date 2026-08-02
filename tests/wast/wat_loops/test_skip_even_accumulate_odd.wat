;; vybe-test: wast/wat_loops/test_skip_even_accumulate_odd
;; origin: languages/wast/tests/wast/test_wat_loops.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        (local $i i32) (local $s i32) i32.const 1 local.set $i
        block loop
          local.get $i i32.const 10 i32.gt_s br_if 1
          local.get $i i32.const 2 i32.rem_u
          if local.get $s local.get $i i32.add local.set $s end
          local.get $i i32.const 1 i32.add local.set $i br 0
        end end local.get $s call $log)
)
