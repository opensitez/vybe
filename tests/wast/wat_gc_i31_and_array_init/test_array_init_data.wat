;; vybe-test: wast/wat_gc_i31_and_array_init/test_array_init_data
;; origin: languages/wast/tests/wast/test_wat_gc_i31_and_array_init.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
  (func $vybe_check_i32 (param i32) (param i32)
    local.get 0
    local.get 1
    i32.ne
    if
      unreachable
    end)
        (type $a (array (mut i32)))
        (data $d "\07\00\00\00\08\00\00\00")
        (func (export "_start") (local $arr (ref $a))
          i32.const 0 i32.const 4 array.new_default $a local.set $arr
          local.get $arr i32.const 0 i32.const 0 i32.const 2 array.init_data $a $d
          local.get $arr i32.const 1 array.get $a i32.const 8 call $vybe_check_i32))
