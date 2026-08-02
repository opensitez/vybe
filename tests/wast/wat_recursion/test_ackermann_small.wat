;; vybe-test: wast/wat_recursion/test_ackermann_small
;; origin: languages/wast/tests/wast/test_wat_recursion.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $ack (param $m i32) (param $n i32) (result i32)
          local.get $m i32.eqz
          if (result i32) local.get $n i32.const 1 i32.add
          else local.get $n i32.eqz
               if (result i32) local.get $m i32.const 1 i32.sub i32.const 1 call $ack
               else local.get $m i32.const 1 i32.sub
                    local.get $m local.get $n i32.const 1 i32.sub call $ack
                    call $ack end end)
        (func (export "_start") i32.const 2 i32.const 3 call $ack call $log))
