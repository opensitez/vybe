use crate::wat_exec;

wat_exec! {
    // `return_call` is a tail call: the caller's result type must match the
    // callee's (WASM rejects a `[]`-result caller tail-calling an `i32`-result
    // func), so `_start` declares `(result i32)`. Its result is the output,
    // the way `wasmtime --invoke` prints an exported function's return value.
    test_return_call_direct => { r#"
(func $f1 (result i32)
  i32.const 42)
(func (export "_start") (result i32)
  return_call $f1
)
"#, "42" },

    test_return_call_with_args => { r#"
(func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(func (export "_start") (result i32)
  i32.const 10
  i32.const 20
  return_call $add
)
"#, "30" },

    test_return_call_recursive => { r#"
(func $fact (param i32 i32) (result i32)
  local.get 0
  i32.const 0
  i32.eq
  if (result i32)
    local.get 1
  else
    local.get 0
    i32.const 1
    i32.sub
    local.get 1
    local.get 0
    i32.mul
    return_call $fact
  end)
(func (export "_start") (result i32)
  i32.const 5
  i32.const 1
  return_call $fact
)
"#, "120" },

    test_return_call_indirect => { r#"
(type $sig (func (result i32)))
(table 1 funcref)
(func $f1 (result i32) i32.const 42)
(elem (i32.const 0) $f1)
(func (export "_start") (result i32)
  i32.const 0
  return_call_indirect (type $sig)
)
"#, "42" },

    test_return_call_indirect_args => { r#"
(type $sig (func (param i32 i32) (result i32)))
(table 1 funcref)
(func $add (param i32 i32) (result i32)
  local.get 0
  local.get 1
  i32.add)
(elem (i32.const 0) $add)
(func (export "_start") (result i32)
  i32.const 10
  i32.const 20
  i32.const 0
  return_call_indirect (type $sig)
)
"#, "30" },

    test_return_call_indirect_oob => { r#"
(type $sig (func (result i32)))
(table 1 funcref)
(func $f1 (result i32) i32.const 42)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 1
  return_call_indirect (type $sig)
)
"#, "trap" },

    test_return_call_indirect_signature_mismatch => { r#"
(type $sig1 (func (result i32)))
(type $sig2 (func (param i32) (result i32)))
(table 1 funcref)
(func $f1 (type $sig2) 
  local.get 0)
(elem (i32.const 0) $f1)
(func (export "_start")
  i32.const 0
  return_call_indirect (type $sig1)
)
"#, "trap" }
}
