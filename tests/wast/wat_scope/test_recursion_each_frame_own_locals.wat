;; vybe-test: wast/wat_scope/test_recursion_each_frame_own_locals
;; origin: languages/wast/tests/wast/test_wat_scope.rs

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $depth (param $n i32) (result i32) (local $marker i32)
          local.get $n local.set $marker
          local.get $n i32.eqz
          if (result i32) i32.const 0
          else local.get $n i32.const 1 i32.sub call $depth
               local.get $marker i32.add end)
        (func (export "_start") i32.const 4 call $depth call $log))
