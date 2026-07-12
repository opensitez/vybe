use crate::wat_exec;

wat_exec! {
    test_global_mut_i32 => { r#"
(global $g (mut i32) (i32.const 42))
(func (export "_start")
  global.get $g
  i32.const 100
  i32.add
  global.set $g
  global.get $g
  call $log
)
"#, "142" },

    test_global_mut_f64 => { r#"
(global $g (mut f64) (f64.const 3.14))
(func (export "_start")
  global.get $g
  f64.const 10.0
  f64.add
  global.set $g
  global.get $g
  i32.trunc_f64_s
  call $log
)
"#, "13" },

    test_global_mut_init_expr => { r#"
(global $a i32 (i32.const 10))
(global $b (mut i32) (global.get $a))
(func (export "_start")
  global.get $b
  i32.const 20
  i32.add
  global.set $b
  global.get $b
  call $log
)
"#, "30" }
}
