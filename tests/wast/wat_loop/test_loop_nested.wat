;; vybe-test: wast/wat_loop/test_loop_nested
;; origin: languages/wast/tests/wast/test_wat_loop.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") (local $i i32) (local $j i32) (local $sum i32)
  i32.const 0
  local.set $i
  block
    loop $outer
      local.get $i
      i32.const 3
      i32.eq
      br_if 1
      
      i32.const 0
      local.set $j
      block
        loop $inner
          local.get $j
          i32.const 2
          i32.eq
          br_if 1
          
          local.get $sum
          i32.const 1
          i32.add
          local.set $sum
          
          local.get $j
          i32.const 1
          i32.add
          local.set $j
          br 0
        end
      end
      
      local.get $i
      i32.const 1
      i32.add
      local.set $i
      br $outer
    end
  end
  local.get $sum
  call $log)
)
