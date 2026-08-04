;; vybe-test: wast/wat_br_table/test_br_table_with_args_default
;; origin: languages/wast/tests/wast/test_wat_br_table.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
  (func (export "_start")
  block (result i32)
    block (result i32)
      block (result i32)
        i32.const 42
        i32.const 99
        br_table 0 1 2
      end
      i32.const 1
      i32.add
      br 1
    end
    i32.const 2
      i32.add
  end
  i32.const 42 call $vybe_check_i32
)
)
