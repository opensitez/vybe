;; vybe-test: wast/wat_try_table/test_try_table_multi_payload
;; origin: languages/wast/tests/wast/test_wat_try_table.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (tag $e (param i32 i32))
(func (export "_start")
  (block $h (result i32 i32)
    (try_table (catch $e $h)
      i32.const 10
      i32.const 20
      throw $e)
    unreachable)
  i32.add
  call $log)
)
