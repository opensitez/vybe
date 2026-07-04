use crate::wat_exec;

wat_exec! {
    test_if_true => { r#"
(func (export "_start")
  i32.const 1
  if
    i32.const 42
    call $log
  end
)
"#, "42" },

    test_if_false => { r#"
(func (export "_start")
  i32.const 0
  if
    i32.const 42
    call $log
  end
  i32.const 99
  call $log
)
"#, "99" },

    test_if_else_true => { r#"
(func (export "_start")
  i32.const 1
  if
    i32.const 42
    call $log
  else
    i32.const 99
    call $log
  end
)
"#, "42" },

    test_if_else_false => { r#"
(func (export "_start")
  i32.const 0
  if
    i32.const 42
    call $log
  else
    i32.const 99
    call $log
  end
)
"#, "99" },

    test_if_result_true => { r#"
(func (export "_start")
  i32.const 1
  if (result i32)
    i32.const 42
  else
    i32.const 99
  end
  call $log
)
"#, "42" },

    test_if_result_false => { r#"
(func (export "_start")
  i32.const 0
  if (result i32)
    i32.const 42
  else
    i32.const 99
  end
  call $log
)
"#, "99" },

    test_if_multi_result => { r#"
(func (export "_start")
  i32.const 1
  if (result i32 i32)
    i32.const 10
    i32.const 20
  else
    i32.const 30
    i32.const 40
  end
  i32.add
  call $log
)
"#, "30" },

    test_if_param_result => { r#"
(func (export "_start")
  i32.const 10
  i32.const 1
  if (param i32) (result i32)
    i32.const 5
    i32.add
  else
    i32.const 2
    i32.mul
  end
  call $log
)
"#, "15" },

    test_if_param_result_false => { r#"
(func (export "_start")
  i32.const 10
  i32.const 0
  if (param i32) (result i32)
    i32.const 5
    i32.add
  else
    i32.const 2
    i32.mul
  end
  call $log
)
"#, "20" },

    test_if_nested => { r#"
(func (export "_start")
  i32.const 1
  if (result i32)
    i32.const 0
    if (result i32)
      i32.const 10
    else
      i32.const 20
    end
  else
    i32.const 30
  end
  call $log
)
"#, "20" },

    test_if_break => { r#"
(func (export "_start")
  block (result i32)
    i32.const 1
    if (result i32)
      i32.const 42
      br 1
    else
      i32.const 99
    end
  end
  call $log
)
"#, "42" },

    test_if_else_break => { r#"
(func (export "_start")
  block (result i32)
    i32.const 0
    if (result i32)
      i32.const 42
    else
      i32.const 99
      br 1
    end
  end
  call $log
)
"#, "99" },

    test_if_negative_condition => { r#"
(func (export "_start")
  i32.const -1
  if (result i32)
    i32.const 42
  else
    i32.const 99
  end
  call $log
)
"#, "42" } // any non-zero is true
}
