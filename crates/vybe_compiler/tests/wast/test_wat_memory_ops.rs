use crate::wat_exec;

wat_exec! {
    test_memory_size => { r#"
(memory 2)
(func (export "_start")
  memory.size
  call $log
)
"#, "2" },

    test_memory_grow => { r#"
(memory 1)
(func (export "_start")
  (i32.const 2)
  memory.grow
  drop
  memory.size
  call $log
)
"#, "3" },

    test_memory_fill => { r#"
(memory 1)
(func (export "_start")
  (i32.const 0) ;; dest
  (i32.const 255) ;; val
  (i32.const 4) ;; len
  memory.fill
  (i32.const 0)
  i32.load
  call $log
)
"#, "-1" },

    test_memory_copy => { r#"
(memory 1)
(func (export "_start")
  (i32.const 0) ;; dest
  (i32.const 255) ;; val
  (i32.const 4) ;; len
  memory.fill
  
  (i32.const 10) ;; dest
  (i32.const 0) ;; src
  (i32.const 4) ;; len
  memory.copy

  (i32.const 10)
  i32.load
  call $log
)
"#, "-1" }
}
