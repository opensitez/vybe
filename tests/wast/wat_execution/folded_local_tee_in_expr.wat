;; vybe-test: wast/wat_execution/folded_local_tee_in_expr
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $x i32)
    (call $log (i32.add (local.tee $x (i32.const 10)) (i32.const 5)))))
