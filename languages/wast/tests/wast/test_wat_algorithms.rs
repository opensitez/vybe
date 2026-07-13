//! Algorithm concepts — complete small programs implemented in WAT and checked
//! against their known results.
use crate::wat_exec;

wat_exec! {
    test_gcd_recursive => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $gcd (param $a i32) (param $b i32) (result i32)
          local.get $b i32.eqz
          if (result i32) local.get $a
          else local.get $b local.get $a local.get $b i32.rem_u call $gcd end)
        (func (export "_start") i32.const 1071 i32.const 462 call $gcd call $log))"#, "21" },
    test_is_prime => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $prime (param $n i32) (result i32) (local $i i32)
          local.get $n i32.const 2 i32.lt_s if i32.const 0 return end
          i32.const 2 local.set $i
          block loop
            local.get $i local.get $i i32.mul local.get $n i32.gt_s br_if 1
            local.get $n local.get $i i32.rem_u i32.eqz if i32.const 0 return end
            local.get $i i32.const 1 i32.add local.set $i br 0
          end end i32.const 1)
        (func (export "_start") i32.const 97 call $prime call $log))"#, "1" },
    test_is_not_prime => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $prime (param $n i32) (result i32) (local $i i32)
          local.get $n i32.const 2 i32.lt_s if i32.const 0 return end
          i32.const 2 local.set $i
          block loop
            local.get $i local.get $i i32.mul local.get $n i32.gt_s br_if 1
            local.get $n local.get $i i32.rem_u i32.eqz if i32.const 0 return end
            local.get $i i32.const 1 i32.add local.set $i br 0
          end end i32.const 1)
        (func (export "_start") i32.const 91 call $prime call $log))"#, "0" },
    test_integer_sqrt_newton => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $isqrt (param $n i32) (result i32) (local $x i32) (local $y i32)
          local.get $n local.set $x
          local.get $n i32.const 1 i32.add i32.const 2 i32.div_u local.set $y
          block loop
            local.get $y local.get $x i32.lt_u i32.eqz br_if 1
            local.get $y local.set $x
            local.get $y local.get $n local.get $y i32.div_u i32.add i32.const 2 i32.div_u local.set $y
            br 0
          end end local.get $x)
        (func (export "_start") i32.const 144 call $isqrt call $log))"#, "12" },
    test_collatz_steps => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $collatz (param $n i32) (result i32) (local $steps i32)
          block loop
            local.get $n i32.const 1 i32.le_u br_if 1
            local.get $n i32.const 1 i32.and
            if local.get $n i32.const 3 i32.mul i32.const 1 i32.add local.set $n
            else local.get $n i32.const 2 i32.div_u local.set $n end
            local.get $steps i32.const 1 i32.add local.set $steps br 0
          end end local.get $steps)
        (func (export "_start") i32.const 6 call $collatz call $log))"#, "8" },
    test_sum_of_squares => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") (local $i i32) (local $s i32) i32.const 1 local.set $i
          block loop local.get $i i32.const 5 i32.gt_s br_if 1
            local.get $s local.get $i local.get $i i32.mul i32.add local.set $s
            local.get $i i32.const 1 i32.add local.set $i br 0 end end
          local.get $s call $log))"#, "55" },
    test_count_set_bits => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const 0xB7 i32.popcnt call $log))"#, "6" },
    test_reverse_digits => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $rev (param $n i32) (result i32) (local $r i32)
          block loop local.get $n i32.eqz br_if 1
            local.get $r i32.const 10 i32.mul local.get $n i32.const 10 i32.rem_u i32.add local.set $r
            local.get $n i32.const 10 i32.div_u local.set $n br 0 end end
          local.get $r)
        (func (export "_start") i32.const 12345 call $rev call $log))"#, "54321" },
    test_power_mod => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $powmod (param $b i32) (param $e i32) (param $m i32) (result i32) (local $r i32)
          i32.const 1 local.set $r
          block loop local.get $e i32.eqz br_if 1
            local.get $r local.get $b i32.mul local.get $m i32.rem_u local.set $r
            local.get $e i32.const 1 i32.sub local.set $e br 0 end end
          local.get $r)
        (func (export "_start") i32.const 2 i32.const 10 i32.const 1000 call $powmod call $log))"#, "24" },
    test_triangular_number => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $tri (param $n i32) (result i32)
          local.get $n local.get $n i32.const 1 i32.add i32.mul i32.const 2 i32.div_u)
        (func (export "_start") i32.const 100 call $tri call $log))"#, "5050" },
}
