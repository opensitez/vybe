//! Recursive function patterns — self recursion, mutual recursion, accumulator
//! style, and deeper call trees.
use crate::wat_exec;

wat_exec! {
    test_tail_recursive_sum => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $sum (param $n i32) (param $acc i32) (result i32)
          local.get $n i32.eqz
          if (result i32) local.get $acc
          else local.get $n i32.const 1 i32.sub
               local.get $acc local.get $n i32.add call $sum end)
        (func (export "_start") i32.const 10 i32.const 0 call $sum call $log))"#, "55" },
    test_power_by_recursion => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $pow (param $b i32) (param $e i32) (result i32)
          local.get $e i32.eqz
          if (result i32) i32.const 1
          else local.get $b local.get $b local.get $e i32.const 1 i32.sub call $pow i32.mul end)
        (func (export "_start") i32.const 2 i32.const 10 call $pow call $log))"#, "1024" },
    test_mutual_recursion_is_even => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $even (param $n i32) (result i32)
          local.get $n i32.eqz
          if (result i32) i32.const 1
          else local.get $n i32.const 1 i32.sub call $odd end)
        (func $odd (param $n i32) (result i32)
          local.get $n i32.eqz
          if (result i32) i32.const 0
          else local.get $n i32.const 1 i32.sub call $even end)
        (func (export "_start") i32.const 10 call $even call $log))"#, "1" },
    test_gcd_euclid => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $gcd (param $a i32) (param $b i32) (result i32)
          local.get $b i32.eqz
          if (result i32) local.get $a
          else local.get $b local.get $a local.get $b i32.rem_u call $gcd end)
        (func (export "_start") i32.const 48 i32.const 36 call $gcd call $log))"#, "12" },
    test_ackermann_small => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $ack (param $m i32) (param $n i32) (result i32)
          local.get $m i32.eqz
          if (result i32) local.get $n i32.const 1 i32.add
          else local.get $n i32.eqz
               if (result i32) local.get $m i32.const 1 i32.sub i32.const 1 call $ack
               else local.get $m i32.const 1 i32.sub
                    local.get $m local.get $n i32.const 1 i32.sub call $ack
                    call $ack end end)
        (func (export "_start") i32.const 2 i32.const 3 call $ack call $log))"#, "9" },
    test_sum_of_digits => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $dig (param $n i32) (result i32)
          local.get $n i32.const 10 i32.lt_u
          if (result i32) local.get $n
          else local.get $n i32.const 10 i32.rem_u
               local.get $n i32.const 10 i32.div_u call $dig i32.add end)
        (func (export "_start") i32.const 12345 call $dig call $log))"#, "15" },
    test_deep_recursion_countdown => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $count (param $n i32) (result i32)
          local.get $n i32.eqz
          if (result i32) i32.const 0
          else local.get $n i32.const 1 i32.sub call $count i32.const 1 i32.add end)
        (func (export "_start") i32.const 500 call $count call $log))"#, "500" },
}
