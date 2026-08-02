;; vybe-test: wast/wat_execution/loop_countdown
;; origin: languages/wast/tests/wast/test_wat_execution.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func (export "_start")
    (local $i i32)
    i32.const 3
    local.set $i
    block $done
      loop $again
        local.get $i
        i32.eqz
        br_if $done
        local.get $i
        call $log
        local.get $i
        i32.const 1
        i32.sub
        local.set $i
        br $again
      end
    end))
