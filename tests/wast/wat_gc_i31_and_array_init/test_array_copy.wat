;; vybe-test: wast/wat_gc_i31_and_array_init/test_array_copy
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
        (func (export "_start") (local $src (ref $a)) (local $dst (ref $a))
          i32.const 55 i32.const 3 array.new $a local.set $src
          i32.const 0 i32.const 3 array.new_default $a local.set $dst
          local.get $dst i32.const 0 local.get $src i32.const 0 i32.const 3 array.copy $a $a
          local.get $dst i32.const 1 array.get $a i32.const 55 call $vybe_check_i32))
