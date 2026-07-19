//! Canonical WASM 3.0 exception handling. `try_table` names, per clause, a tag
//! and a branch LABEL a matching thrown exception transfers to, delivering the
//! tag payload — and, for `_ref` clauses, the caught `exnref`. `throw` raises;
//! `throw_ref` re-raises an `exnref`. Matching is by TAG IDENTITY only. These
//! replace the deprecated legacy `try/catch/catch_all/delegate/rethrow` text
//! surface — the VM has only this mechanism.
use crate::wat_exec;

wat_exec! {
    // `catch $e $h` transfers to $h delivering the tag's i32 payload (42).
    test_try_table_catch_payload => { r#"
(tag $e (param i32))
(func (export "_start")
  (block $h (result i32)
    (try_table (catch $e $h)
      i32.const 42
      throw $e)
    i32.const 0)
  call $log)
"#, "42" },

    // No exception thrown → try_table completes normally with its body's result.
    test_try_table_no_throw => { r#"
(func (export "_start")
  (try_table (result i32)
    i32.const 99)
  call $log)
"#, "99" },

    // `catch_all` matches any tag and branches with no payload; the handler code
    // (after the guarding block) supplies the value.
    test_try_table_catch_all => { r#"
(tag $e (param i32))
(func (export "_start")
  (block $h
    (try_table (catch_all $h)
      i32.const 42
      throw $e)
    unreachable)
  i32.const 99
  call $log)
"#, "99" },

    // A two-value payload is delivered to a multi-result target block.
    test_try_table_multi_payload => { r#"
(tag $e (param i32 i32))
(func (export "_start")
  (block $h (result i32 i32)
    (try_table (catch $e $h)
      i32.const 10
      i32.const 20
      throw $e)
    unreachable)
  i32.add
  call $log)
"#, "30" },

    // Tag identity: the inner clause names $e1, so a thrown $e2 skips it and is
    // caught by the outer clause. (Never name/subtype matching.)
    test_try_table_nested_tag_identity => { r#"
(tag $e1 (param i32))
(tag $e2 (param i32))
(func (export "_start")
  (block $outer (result i32)
    (try_table (catch $e2 $outer)
      (try_table (catch $e1 $outer)
        i32.const 42
        throw $e2)
      unreachable)
    i32.const 0)
  call $log)
"#, "42" },

    // `catch_ref` delivers payload + the caught `exnref`; `throw_ref` re-raises
    // that exnref, which the outer clause then catches for its payload.
    test_try_table_catch_ref_throw_ref => { r#"
(tag $e (param i32))
(func (export "_start")
  (block $outer (result i32)
    (try_table (catch $e $outer)
      (block $h (result exnref)
        (try_table (catch_all_ref $h)
          i32.const 42
          throw $e)
        unreachable)
      throw_ref)
    i32.const 0)
  call $log)
"#, "42" },

    // An exception raised in a callee unwinds across the call frame to the
    // caller's try_table (tags are one load-time entity shared across chunks).
    test_try_table_propagates_across_call => { r#"
(tag $e (param i32))
(func $raise
  i32.const 7
  throw $e)
(func (export "_start")
  (block $h (result i32)
    (try_table (catch $e $h)
      call $raise
      i32.const 0)
    i32.const 0)
  call $log)
"#, "7" },
}
