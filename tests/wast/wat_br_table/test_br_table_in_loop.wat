;; vybe-test: wast/wat_br_table/test_br_table_in_loop
;; origin: languages/wast/tests/wast/test_wat_br_table.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") (local $i i32)
  i32.const 0
  local.set $i
  block
    loop
      local.get $i
      i32.const 1
      i32.add
      local.set $i
      
      local.get $i
      br_table 1 0 1 ;; if $i==0 break outer, if $i==1 continue loop, if $i>=2 break outer
    end
  end
  local.get $i
  call $log
)
)
