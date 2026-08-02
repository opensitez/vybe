;; vybe-test: wast/wat_scope/test_inner_label_shadows_by_depth
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
        block block
          i32.const 7 call $log br 0 i32.const 8 call $log
        end i32.const 9 call $log end)
)
