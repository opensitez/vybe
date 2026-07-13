//! Branch hinting proposal — `@metadata.code.branch_hint` annotations on
//! br_if / if. They are hints only; behaviour is unchanged, so these verify the
//! annotated forms parse and run to the same result.
use crate::wat_exec;

wat_exec! {
    test_branch_hint_likely_taken => { r#"(func (export "_start")
        (@metadata.code.branch_hint "\01") i32.const 1
        if (result i32) i32.const 42 else i32.const 0 end call $log)"#, "42" },
    test_branch_hint_unlikely => { r#"(func (export "_start")
        (@metadata.code.branch_hint "\00") i32.const 0
        if (result i32) i32.const 1 else i32.const 99 end call $log)"#, "99" },
    test_branch_hint_on_br_if => { r#"(func (export "_start")
        block i32.const 7 call $log
          (@metadata.code.branch_hint "\01") i32.const 1 br_if 0
          i32.const 8 call $log
        end)"#, "7" },
}
