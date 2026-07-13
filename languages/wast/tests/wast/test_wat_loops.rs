//! Loop concepts — iteration patterns expressed with block/loop/br/br_if:
//! counting, accumulation, early exit, nested iteration, and loop-carried state.
use crate::wat_exec;

wat_exec! {
    test_count_up_to_n => { r#"(func (export "_start")
        (local $i i32) (local $c i32)
        block loop
          local.get $i i32.const 10 i32.ge_s br_if 1
          local.get $c i32.const 1 i32.add local.set $c
          local.get $i i32.const 1 i32.add local.set $i br 0
        end end local.get $c call $log)"#, "10" },
    test_count_down_to_zero => { r#"(func (export "_start")
        (local $i i32) i32.const 7 local.set $i
        block loop local.get $i i32.eqz br_if 1
          local.get $i i32.const 1 i32.sub local.set $i br 0 end end
        local.get $i call $log)"#, "0" },
    test_sum_1_to_100 => { r#"(func (export "_start")
        (local $i i32) (local $s i32) i32.const 1 local.set $i
        block loop local.get $i i32.const 100 i32.gt_s br_if 1
          local.get $s local.get $i i32.add local.set $s
          local.get $i i32.const 1 i32.add local.set $i br 0 end end
        local.get $s call $log)"#, "5050" },
    test_factorial_via_loop => { r#"(func (export "_start")
        (local $i i32) (local $f i32) i32.const 1 local.set $i i32.const 1 local.set $f
        block loop local.get $i i32.const 6 i32.gt_s br_if 1
          local.get $f local.get $i i32.mul local.set $f
          local.get $i i32.const 1 i32.add local.set $i br 0 end end
        local.get $f call $log)"#, "720" },
    test_power_of_two_via_loop => { r#"(func (export "_start")
        (local $i i32) (local $p i32) i32.const 1 local.set $p
        block loop local.get $i i32.const 8 i32.ge_s br_if 1
          local.get $p i32.const 2 i32.mul local.set $p
          local.get $i i32.const 1 i32.add local.set $i br 0 end end
        local.get $p call $log)"#, "256" },
    test_early_exit_on_condition => { r#"(func (export "_start")
        (local $i i32)
        block loop
          local.get $i i32.const 5 i32.eq br_if 1
          local.get $i i32.const 1 i32.add local.set $i br 0
        end end local.get $i call $log)"#, "5" },
    test_nested_loop_multiplication_table_cell => { r#"(func (export "_start")
        (local $i i32) (local $j i32) (local $sum i32)
        block loop
          local.get $i i32.const 3 i32.ge_s br_if 1
          i32.const 0 local.set $j
          block loop
            local.get $j i32.const 3 i32.ge_s br_if 1
            local.get $sum local.get $i local.get $j i32.mul i32.add local.set $sum
            local.get $j i32.const 1 i32.add local.set $j br 0
          end end
          local.get $i i32.const 1 i32.add local.set $i br 0
        end end local.get $sum call $log)"#, "9" },
    test_skip_even_accumulate_odd => { r#"(func (export "_start")
        (local $i i32) (local $s i32) i32.const 1 local.set $i
        block loop
          local.get $i i32.const 10 i32.gt_s br_if 1
          local.get $i i32.const 2 i32.rem_u
          if local.get $s local.get $i i32.add local.set $s end
          local.get $i i32.const 1 i32.add local.set $i br 0
        end end local.get $s call $log)"#, "25" },
    test_loop_with_continue_pattern => { r#"(func (export "_start")
        (local $i i32) (local $s i32)
        block loop
          local.get $i i32.const 5 i32.ge_s br_if 1
          local.get $i i32.const 1 i32.add local.set $i
          local.get $i i32.const 3 i32.eq
          if br 1 end
          local.get $s local.get $i i32.add local.set $s
          br 0
        end end local.get $s call $log)"#, "12" },
    test_geometric_series_sum => { r#"(func (export "_start")
        (local $term i32) (local $sum i32) (local $n i32)
        i32.const 1 local.set $term
        block loop
          local.get $n i32.const 5 i32.ge_s br_if 1
          local.get $sum local.get $term i32.add local.set $sum
          local.get $term i32.const 3 i32.mul local.set $term
          local.get $n i32.const 1 i32.add local.set $n br 0
        end end local.get $sum call $log)"#, "121" },
    test_fibonacci_iterative => { r#"(func (export "_start")
        (local $a i32) (local $b i32) (local $i i32) (local $t i32)
        i32.const 0 local.set $a i32.const 1 local.set $b
        block loop
          local.get $i i32.const 10 i32.ge_s br_if 1
          local.get $a local.get $b i32.add local.set $t
          local.get $b local.set $a local.get $t local.set $b
          local.get $i i32.const 1 i32.add local.set $i br 0
        end end local.get $a call $log)"#, "55" },
    test_digit_count => { r#"(func (export "_start")
        (local $n i32) (local $c i32) i32.const 12345 local.set $n
        block loop
          local.get $n i32.eqz br_if 1
          local.get $n i32.const 10 i32.div_u local.set $n
          local.get $c i32.const 1 i32.add local.set $c br 0
        end end local.get $c call $log)"#, "5" },
}
