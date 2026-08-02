;; vybe-test: wast/wat_try_table/test_try_table_catch_ref_throw_ref
;; origin: languages/wast/tests/wast/test_wat_try_table.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (tag $e (param i32))
(func (export "_start")
  (block $outer (result i32)
    (try_table (catch $e $outer)
      (block $h (result exnref)
        (try_table (catch_all_ref $h)
          i32.const 42
          throw $e)
        unreachable)
      throw_ref)
    i32.const 0)
  call $log)
)
