use crate::wat_exec;

wat_exec! {
    test_br_table_first => { r#"
(func (export "_start")
  block (result i32)
    block
      block
        i32.const 0
        br_table 0 1 2
      end
      i32.const 10
      br 1
    end
    i32.const 20
  end
  call $log
)
"#, "10" },

    test_br_table_second => { r#"
(func (export "_start")
  block (result i32)
    block
      block
        i32.const 1
        br_table 0 1 2
      end
      i32.const 10
      br 1
    end
    i32.const 20
  end
  call $log
)
"#, "20" },

    // Selector == table length picks the default (last) label. All br_table
    // targets must share result arity, so every block yields i32 (a mismatched
    // module is invalid WASM — wasmtime rejects it). Default -> outer block
    // carries 42 straight out, skipping the +1/+2 fall-through paths.
    test_br_table_default => { r#"
(func (export "_start")
  block (result i32)
    block (result i32)
      block (result i32)
        i32.const 42
        i32.const 2
        br_table 0 1 2
      end
      i32.const 1
      i32.add
      br 1
    end
    i32.const 2
    i32.add
  end
  call $log
)
"#, "42" },

    test_br_table_default_out_of_bounds => { r#"
(func (export "_start")
  block (result i32)
    block (result i32)
      block (result i32)
        i32.const 42
        i32.const 99
        br_table 0 1 2
      end
      i32.const 1
      i32.add
      br 1
    end
    i32.const 2
    i32.add
  end
  call $log
)
"#, "42" }, // out-of-range selector also takes the default label

    test_br_table_with_args_first => { r#"
(func (export "_start")
  block (result i32)
    block (result i32)
      block (result i32)
        i32.const 42
        i32.const 0
        br_table 0 1 2
      end
      i32.const 1
      i32.add
      br 1
    end
    i32.const 2
      i32.add
  end
  call $log
)
"#, "43" }, // 42 + 1

    test_br_table_with_args_second => { r#"
(func (export "_start")
  block (result i32)
    block (result i32)
      block (result i32)
        i32.const 42
        i32.const 1
        br_table 0 1 2
      end
      i32.const 1
      i32.add
      br 1
    end
    i32.const 2
      i32.add
  end
  call $log
)
"#, "44" }, // 42 + 2

    test_br_table_with_args_default => { r#"
(func (export "_start")
  block (result i32)
    block (result i32)
      block (result i32)
        i32.const 42
        i32.const 99
        br_table 0 1 2
      end
      i32.const 1
      i32.add
      br 1
    end
    i32.const 2
      i32.add
  end
  call $log
)
"#, "42" }, // breaks out of the outermost block, returns 42

    test_br_table_in_loop => { r#"
(func (export "_start") (local $i i32)
  i32.const 0
  local.set $i
  block
    loop
      local.get $i
      i32.const 1
      i32.add
      local.set $i
      
      local.get $i
      br_table 1 0 1 ;; if $i==0 break outer, if $i==1 continue loop, if $i>=2 break outer
    end
  end
  local.get $i
  call $log
)
"#, "2" },

    test_br_table_empty => { r#"
(func (export "_start")
  block (result i32)
    i32.const 10
    i32.const 0
    br_table 0
  end
  call $log
)
"#, "10" } // only default target
}
