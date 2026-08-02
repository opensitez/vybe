;; vybe-test: wast/wat_control_flow/compile_loop_sum
;; origin: languages/wast/tests/wast/test_wat_control_flow.rs
;; vybe-test-mode: compile

(module
  (func $sum (export "sum") (param $n i32) (result i32)
    (local $acc i32)
    (local $i i32)
    i32.const 0
    local.set $acc
    i32.const 1
    local.set $i
    (block $break
      (loop $continue
        local.get $i
        local.get $n
        i32.gt_s
        br_if $break
        local.get $acc
        local.get $i
        i32.add
        local.set $acc
        local.get $i
        i32.const 1
        i32.add
        local.set $i
        br $continue))
    local.get $acc))
