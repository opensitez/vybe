;; vybe-test: wast/wat_loops/test_digit_count
;; origin: languages/wast/tests/wast/test_wat_loops.rs

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
        (local $n i32) (local $c i32) i32.const 12345 local.set $n
        block loop
          local.get $n i32.eqz br_if 1
          local.get $n i32.const 10 i32.div_u local.set $n
          local.get $c i32.const 1 i32.add local.set $c br 0
        end end local.get $c i32.const 5 call $vybe_check_i32)
)
