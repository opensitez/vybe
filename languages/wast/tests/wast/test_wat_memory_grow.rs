use crate::wat_exec;

wat_exec! {
    test_memory_grow_success => { r#"
(memory 1)
(func (export "_start") 
  i32.const 2
  memory.grow
  call $log)
"#, "1" }, // returns old size

    test_memory_grow_multiple => { r#"
(memory 1)
(func (export "_start") 
  i32.const 2
  memory.grow
  drop
  i32.const 3
  memory.grow
  call $log)
"#, "3" }, // returns old size (1 + 2 = 3)

    test_memory_grow_zero => { r#"
(memory 1)
(func (export "_start") 
  i32.const 0
  memory.grow
  call $log)
"#, "1" },

    test_memory_grow_fail_max => { r#"
(memory 1 2)
(func (export "_start") 
  i32.const 5
  memory.grow
  call $log)
"#, "-1" }, // fails because 1+5 > 2

    test_memory_size_initial => { r#"
(memory 5)
(func (export "_start") 
  memory.size
  call $log)
"#, "5" },

    test_memory_size_after_grow => { r#"
(memory 1)
(func (export "_start") 
  i32.const 4
  memory.grow
  drop
  memory.size
  call $log)
"#, "5" },
    
    test_memory_size_after_failed_grow => { r#"
(memory 1 2)
(func (export "_start") 
  i32.const 5
  memory.grow
  drop
  memory.size
  call $log)
"#, "1" }
}
