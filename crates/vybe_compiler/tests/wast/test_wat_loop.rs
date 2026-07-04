use crate::wat_exec;

wat_exec! {
    test_loop_empty => { r#"
(func (export "_start")
  loop
  end
  i32.const 42
  call $log)
"#, "42" },

    test_loop_value => { r#"
(func (export "_start")
  loop (result i32)
    i32.const 10
  end
  call $log)
"#, "10" },

    test_loop_break_out => { r#"
(func (export "_start")
  block (result i32)
    loop
      i32.const 42
      br 1
    end
  end
  call $log)
"#, "42" },

    test_loop_continue => { r#"
(func (export "_start") (local $i i32)
  i32.const 0
  local.set $i
  block
    loop
      local.get $i
      i32.const 5
      i32.eq
      br_if 1
      local.get $i
      i32.const 1
      i32.add
      local.set $i
      br 0
    end
  end
  local.get $i
  call $log)
"#, "5" },

    test_loop_continue_with_params => { r#"
(func (export "_start")
  i32.const 0
  block (result i32)
    loop (param i32) (result i32)
      local.tee 0
      i32.const 5
      i32.eq
      br_if 1
      local.get 0
      i32.const 1
      i32.add
      br 0
    end
  end
  call $log)
"#, "5" },

    test_loop_nested => { r#"
(func (export "_start") (local $i i32) (local $j i32) (local $sum i32)
  i32.const 0
  local.set $i
  block
    loop $outer
      local.get $i
      i32.const 3
      i32.eq
      br_if 1
      
      i32.const 0
      local.set $j
      block
        loop $inner
          local.get $j
          i32.const 2
          i32.eq
          br_if 1
          
          local.get $sum
          i32.const 1
          i32.add
          local.set $sum
          
          local.get $j
          i32.const 1
          i32.add
          local.set $j
          br 0
        end
      end
      
      local.get $i
      i32.const 1
      i32.add
      local.set $i
      br $outer
    end
  end
  local.get $sum
  call $log)
"#, "6" }, // 3 * 2 = 6

    test_loop_unreachable => { r#"
(func (export "_start")
  block (result i32)
    loop (result i32)
      i32.const 42
      br 1
      unreachable
    end
  end
  call $log)
"#, "42" },

    test_loop_no_yield_from_continue => { r#"
(func (export "_start")
  i32.const 0
  block (result i32)
    loop (param i32) (result i32)
      drop
      i32.const 10
      br 1
      i32.const 99
      br 0
    end
  end
  call $log)
"#, "10" }
}
