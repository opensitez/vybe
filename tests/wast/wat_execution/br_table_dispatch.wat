;; vybe-test: wast/wat_execution/br_table_dispatch
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $dispatch (param $x i32) (result i32)
    block $default (result i32)
      block $c2 (result i32)
        block $c1 (result i32)
          block $c0 (result i32)
            local.get $x
            br_table $c0 $c1 $c2 $default
          end
          i32.const 100
          br $default
        end
        i32.const 200
        br $default
      end
      i32.const 300
      br $default
    end)
  (func (export "_start")
    i32.const 0 call $dispatch call $log
    i32.const 1 call $dispatch call $log
    i32.const 2 call $dispatch call $log))
