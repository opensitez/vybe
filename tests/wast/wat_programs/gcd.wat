;; vybe-test: wast/wat_programs/gcd
;; origin: languages/wast/tests/wast/test_wat_programs.rs
;; vybe-test-mode: compile

(module
  (func $gcd (export "gcd") (param $a i32) (param $b i32) (result i32)
    (block $done (result i32)
      (loop $loop
        local.get $b
        i32.eqz
        br_if $done
        local.get $a
        local.get $b
        i32.rem_u
        local.get $b
        local.set $a
        local.set $b
        br $loop)
      local.get $a))
)
