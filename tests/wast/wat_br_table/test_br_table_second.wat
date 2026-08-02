;; vybe-test: wast/wat_br_table/test_br_table_second
;; origin: languages/wast/tests/wast/test_wat_br_table.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start")
  block (result i32)
    block
      block
        i32.const 1
        br_table 0 1 2
      end
      i32.const 10
      br 1
    end
    i32.const 20
  end
  call $log
)
)
