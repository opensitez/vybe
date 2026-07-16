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

    // ── Tail-call frame reuse (WASM tail-call proposal) ──────────────────
    // Depth 300 is past the VM's 256-frame limit, so these three tests form a
    // control: the NON-tail recursion overflows (proving the limit is real),
    // while `return_call`/`return_call_indirect` complete because they reuse
    // the frame. The positive tests return the accumulated sum 300+299+…+1 =
    // 45150 (not a constant), so a shortcut that skips iterations can't pass.

    // Negative control: plain `call` + `return` (work after the call, so NOT a
    // tail call) must exhaust the call stack at this depth.
    test_deep_non_tail_recursion_overflows => { r#"
(func $sum (param $n i32) (result i32)
  local.get $n
  i32.eqz
  if (result i32)
    i32.const 0
  else
    local.get $n
    local.get $n
    i32.const 1
    i32.sub
    call $sum
    i32.add
  end)
(func (export "_start") (result i32)
  i32.const 300
  call $sum
)
"#, "trap" },

    test_return_call_deep_tail_recursion => { r#"
(func $sum (param $n i32) (param $acc i32) (result i32)
  local.get $n
  i32.eqz
  if (result i32)
    local.get $acc
  else
    local.get $n
    i32.const 1
    i32.sub
    local.get $acc
    local.get $n
    i32.add
    return_call $sum
  end)
(func (export "_start") (result i32)
  i32.const 300
  i32.const 0
  return_call $sum
)
"#, "45150" },

    test_return_call_indirect_deep_tail_recursion => { r#"
(type $sig (func (param i32 i32) (result i32)))
(table 1 funcref)
(func $sum (param $n i32) (param $acc i32) (result i32)
  local.get $n
  i32.eqz
  if (result i32)
    local.get $acc
  else
    local.get $n
    i32.const 1
    i32.sub
    local.get $acc
    local.get $n
    i32.add
    i32.const 0
    return_call_indirect (type $sig)
  end)
(elem (i32.const 0) $sum)
(func (export "_start") (result i32)
  i32.const 300  ;; n
  i32.const 0    ;; acc
  i32.const 0    ;; table index
  return_call_indirect (type $sig)
)
"#, "45150" },

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
