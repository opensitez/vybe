;; vybe-test: wast/wat_scope/test_nested_block_labels_distinct
;; origin: languages/wast/tests/wast/test_wat_scope.rs

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
        (local $r i32)
        block $outer block $inner
          i32.const 1 br_if $inner
          i32.const 0 local.set $r br $outer
        end i32.const 42 local.set $r end
        local.get $r i32.const 42 call $vybe_check_i32)
)
