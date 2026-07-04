use crate::wat_exec;

wat_exec! {
    test_global_const_i32 => { r#"
(global $g i32 (i32.const 42))
(func (export "_start")
  global.get $g
  call $log
)
"#, "42" },

    test_global_const_f64 => { r#"
(global $g f64 (f64.const 3.14))
(func (export "_start")
  global.get $g
  i32.trunc_f64_s
  call $log
)
"#, "3" },

    test_global_const_multiple => { r#"
(global $a i32 (i32.const 10))
(global $b i32 (i32.const 20))
(func (export "_start")
  global.get $a
  global.get $b
  i32.add
  call $log
)
"#, "30" }
}
