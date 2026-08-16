;; vybe-test: wast/wast_script_module/module_export_import_chain
;; origin: languages/wast/tests/wast/test_wast_script_module.rs
;; vybe-test-mode: run

(module $m1 (func (export "f") (result i32) i32.const 42))
(register "lib" $m1)
(module $m2 
  (import "lib" "f" (func $f (result i32)))
  (func (export "g") (result i32) call $f i32.const 1 i32.add)
)
(assert_return (invoke $m2 "g") (i32.const 43))
