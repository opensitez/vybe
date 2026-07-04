use crate::wat_exec;

wat_exec! {
    test_block_empty => { r#"
(func (export "_start")
  block
  end
  i32.const 42
  call $log)
"#, "42" },

    test_block_value => { r#"
(func (export "_start")
  block (result i32)
    i32.const 10
  end
  call $log)
"#, "10" },

    test_block_multi_value => { r#"
(func (export "_start")
  block (result i32 i32)
    i32.const 10
    i32.const 20
  end
  i32.add
  call $log)
"#, "30" },

    test_block_break => { r#"
(func (export "_start")
  block (result i32)
    i32.const 10
    br 0
    i32.const 20
  end
  call $log)
"#, "10" },

    test_block_break_multi_value => { r#"
(func (export "_start")
  block (result i32 i32)
    i32.const 10
    i32.const 20
    br 0
    i32.const 30
    i32.const 40
  end
  i32.add
  call $log)
"#, "30" },

    test_block_nested_break_inner => { r#"
(func (export "_start")
  block (result i32)
    i32.const 10
    block
      br 0
      i32.const 50
      drop
    end
    i32.const 20
    i32.add
  end
  call $log)
"#, "30" },

    test_block_nested_break_outer => { r#"
(func (export "_start")
  block (result i32)
    i32.const 10
    block
      i32.const 50
      br 1
    end
    i32.const 20
    i32.add
  end
  call $log)
"#, "50" },

    test_block_nested_break_outer_with_args => { r#"
(func (export "_start")
  block (result i32)
    i32.const 10
    block (param i32)
      i32.const 20
      i32.add
      br 1
    end
    i32.const 50
    i32.add
  end
  call $log)
"#, "30" },

    test_block_params => { r#"
(func (export "_start")
  i32.const 5
  block (param i32) (result i32)
    i32.const 10
    i32.add
  end
  call $log)
"#, "15" },

    test_block_params_break => { r#"
(func (export "_start")
  i32.const 5
  block (param i32) (result i32)
    i32.const 10
    i32.add
    br 0
  end
  call $log)
"#, "15" },

    test_block_deep_nesting => { r#"
(func (export "_start")
  block (result i32)
    block (result i32)
      block (result i32)
        block (result i32)
          i32.const 42
          br 3
        end
      end
    end
  end
  call $log)
"#, "42" },

    test_block_drop_inner => { r#"
(func (export "_start")
  block (result i32)
    i32.const 1
    block (result i32)
      i32.const 2
    end
    drop
  end
  call $log)
"#, "1" }
}
