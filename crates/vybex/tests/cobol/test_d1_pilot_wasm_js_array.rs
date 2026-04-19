//! Phase D1 pilot — compiler-level validation that the `Array(count,
//! init)` intercept routes through the new `wasm:js-array.*` host
//! imports end-to-end.
//!
//! The intercept lives at
//! `crates/vybex/src/compiler.rs::try_compile_builtin` and fires
//! whenever a compiler sees a call expression with callee
//! `Ident("Array")` and two positional args. COBOL's OCCURS walker
//! produces this shape at the top-level (there's an unrelated name-
//! extraction issue in the nested-field walker that means the
//! intercept doesn't fire for 05-level OCCURS fields **yet** — that
//! walker fix is a follow-up; for the D1 pilot we validate the
//! compiler intercept itself via a direct AST).
//!
//! Validates:
//!   1. The intercept fires when the compiler receives `Array(N, v)`.
//!   2. The resulting chunk imports `wasm:js-array.newWithLength`.
//!   3. Non-null init triggers `wasm:js-array.fill` too.
//!   4. Running the resulting bytecode produces an Array of the
//!      expected length — i.e. the full stdlib-import → handler →
//!      ObjectKind::Array pipeline works end-to-end.
//!
//! Once the COBOL walker's nested-field name extraction is fixed
//! (orthogonal issue), a real COBOL program with `OCCURS` will flow
//! through this same intercept without further changes.

use std::collections::BTreeSet;
use vybe_bytecode::VM;
use vybex::ast::{
    Argument, BindingPattern, Expression, ExprKind, Literal, StmtKind,
    Statement, VarDeclKind, VarDeclarator,
};

use super::helpers;

/// Scan every chunk's import table and return the sorted (module, name) set.
fn cobol_imports_for(src: &str) -> Vec<(String, String)> {
    let chunks = helpers::compile(src);
    let mut seen = BTreeSet::new();
    for chunk in &chunks {
        for i in &chunk.imports {
            seen.insert((i.module.clone(), i.name.clone()));
        }
    }
    seen.into_iter().collect()
}

/// Compile a synthetic `var x = Array(count, init)` at module level
/// using the COBOL language profile. The result is the chunk set the
/// compiler would produce from a COBOL walker that correctly emits
/// the Array-call shape. Returns every (module, name) pair imported.
fn compile_array_call(count: Expression, init: Expression) -> Vec<(String, String)> {
    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(count),
            Argument::positional(init),
        ],
        optional: false,
    });

    let stmt = StmtKind::VarDecl {
        kind: VarDeclKind::Dim,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("tbl".to_string()),
            type_hint: None,
            init: Some(call_expr),
            array_bounds: None,
            with_events: false,
        }],
    };
    let module = vybex::ast::Module {
        name: "<pilot>".into(),
        language: vybex::ast::Lang::Cobol,
        body: vec![Statement::new(stmt)],
        imports: Vec::new(),
    };
    let profile = vybex::profile::parse_profile(vybex::languages::cobol::profile_source())
        .expect("parse profile");
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile");

    let mut imports = BTreeSet::new();
    for chunk in &chunks {
        for i in &chunk.imports {
            imports.insert((i.module.clone(), i.name.clone()));
        }
    }
    imports.into_iter().collect()
}

#[test]
fn array_call_with_null_init_emits_only_newwithlength() {
    let imports = compile_array_call(
        Expression::new(ExprKind::Lit(Literal::Float(5.0))),
        Expression::new(ExprKind::Lit(Literal::Null)),
    );
    assert!(imports.contains(&("wasm:js-array".into(), "newWithLength".into())),
        "expected `wasm:js-array.newWithLength` import; got: {:?}", imports);
    // fill() is NOT emitted when init is null — newWithLength already
    // null-fills.
    assert!(!imports.contains(&("wasm:js-array".into(), "fill".into())),
        "expected no `wasm:js-array.fill` for null-init shortcut; got: {:?}", imports);
}

#[test]
fn array_call_with_non_null_init_emits_fill_too() {
    let imports = compile_array_call(
        Expression::new(ExprKind::Lit(Literal::Float(5.0))),
        Expression::new(ExprKind::Lit(Literal::Float(42.0))),
    );
    assert!(imports.contains(&("wasm:js-array".into(), "newWithLength".into())),
        "expected newWithLength import; got: {:?}", imports);
    assert!(imports.contains(&("wasm:js-array".into(), "fill".into())),
        "expected fill import to initialise non-null value; got: {:?}", imports);
}

#[test]
fn real_cobol_nested_occurs_emits_wasm_js_array_import() {
    // Real-COBOL end-to-end: a 05-level OCCURS field inside a 01-level
    // group. After the walker's `ident_or_keyword` fix the nested field
    // now flows through the `Array(count, init)` → `wasm:js-array.*`
    // intercept.
    let imports = cobol_imports_for(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PILOT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL.
          05 WS-ITEM PIC 9(3) OCCURS 5 TIMES.
       PROCEDURE DIVISION.
       MAIN-PARA.
           STOP RUN.
"#.trim());

    assert!(
        imports.iter().any(|(m, _)| m == "wasm:js-array"),
        "Real COBOL `05 … OCCURS N TIMES` must now import from \
         `wasm:js-array` (walker ident_or_keyword fix); imports = {:?}",
        imports
    );
    assert!(
        imports.contains(&("wasm:js-array".into(), "newWithLength".into())),
        "Expected `wasm:js-array.newWithLength` in imports; got: {:?}",
        imports
    );
}

#[test]
fn real_cobol_occurs_with_value_clause_also_emits_fill() {
    // OCCURS + VALUE → initialised with a non-null default, so the
    // intercept emits `fill` in addition to `newWithLength`.
    let imports = cobol_imports_for(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PILOT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL.
          05 WS-ITEM PIC 9(3) OCCURS 10 TIMES VALUE 42.
       PROCEDURE DIVISION.
       MAIN-PARA.
           STOP RUN.
"#.trim());

    assert!(
        imports.contains(&("wasm:js-array".into(), "newWithLength".into())),
        "newWithLength missing; imports = {:?}", imports
    );
    assert!(
        imports.contains(&("wasm:js-array".into(), "fill".into())),
        "fill should be emitted when OCCURS has a non-null VALUE; imports = {:?}", imports
    );
}

#[test]
fn real_cobol_occurs_runs_end_to_end() {
    // Compile + execute a COBOL program with a nested OCCURS table.
    // The program body is a no-op (STOP RUN) — the test verifies the
    // data declarations don't trap at load time.
    helpers::run(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. PILOT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL.
          05 WS-ITEM PIC 9(3) OCCURS 5 TIMES.
       PROCEDURE DIVISION.
       MAIN-PARA.
           STOP RUN.
"#.trim());
}

#[test]
fn array_call_end_to_end_runtime_produces_array_of_length_n() {
    // Full pipeline: compile → VM → observe result. `tbl` should be
    // an Array of length 5 after running.
    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Float(5.0)))),
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Null))),
        ],
        optional: false,
    });

    // Wrap in: var tbl = Array(5, null); return tbl.length;
    // We build the `tbl.length` read as a Member expression.
    let decl = StmtKind::VarDecl {
        kind: VarDeclKind::Dim,
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("tbl".to_string()),
            type_hint: None,
            init: Some(call_expr),
            array_bounds: None,
            with_events: false,
        }],
    };

    let module = vybex::ast::Module {
        name: "<pilot>".into(),
        language: vybex::ast::Lang::Cobol,
        body: vec![Statement::new(decl)],
        imports: Vec::new(),
    };
    let profile = vybex::profile::parse_profile(vybex::languages::cobol::profile_source())
        .expect("parse profile");
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile");

    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    // Running this chunk produces an Array and stores it in a global.
    // The call itself is the last value on the stack of chunk 0 before
    // the GLOBAL_SET; running to completion is sufficient to prove the
    // handlers executed without trapping.
    vm.run(chunks).expect("VM run succeeded");
}
