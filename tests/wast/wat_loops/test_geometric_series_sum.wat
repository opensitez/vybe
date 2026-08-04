;; vybe-test: wast/wat_loops/test_geometric_series_sum
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
        (local $term i32) (local $sum i32) (local $n i32)
        i32.const 1 local.set $term
        block loop
          local.get $n i32.const 5 i32.ge_s br_if 1
          local.get $sum local.get $term i32.add local.set $sum
          local.get $term i32.const 3 i32.mul local.set $term
          local.get $n i32.const 1 i32.add local.set $n br 0
        end end local.get $sum i32.const 121 call $vybe_check_i32)
)
