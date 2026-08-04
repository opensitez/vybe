;; vybe-test: wast/wat_block/test_block_drop_inner
;; origin: languages/wast/tests/wast/test_wat_block.rs

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
    i32.const 1
    block (result i32)
      i32.const 2
    end
    drop
  end
  i32.const 1 call $vybe_check_i32)
)
