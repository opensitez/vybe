use crate::helpers::{run_in_main, run_main};

#[test]
fn throw_new_runtime_exception_without_message() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"thrown\"); }",
    );
    assert_eq!(out, vec!["thrown"]);
}

#[test]
fn throw_new_runtime_exception_with_literal_message() {
    let out = run_main(
        "try { throw new RuntimeException(\"abort\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["abort"]);
}

#[test]
fn throw_new_illegal_argument_with_expression_message() {
    let out = run_main(
        "try { throw new IllegalArgumentException(\"limit:\" + 10); } catch (IllegalArgumentException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["limit:10"]);
}

#[test]
fn throw_new_null_pointer_exception_reaches_catch() {
    let out = run_main(
        "try { throw new NullPointerException(\"null ref\"); } catch (NullPointerException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["null ref"]);
}

#[test]
fn throw_new_index_out_of_bounds_reaches_catch() {
    let out = run_main(
        "try { throw new IndexOutOfBoundsException(\"bad index\"); } catch (IndexOutOfBoundsException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["bad index"]);
}

#[test]
fn throw_new_io_exception_reaches_catch() {
    let out = run_main(
        "try { throw new java.io.IOException(\"broken pipe\"); } catch (java.io.IOException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["broken pipe"]);
}

#[test]
fn rethrow_preserves_original_message_to_outer_handler() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"deep\"); } catch (RuntimeException e) { throw e; } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["deep"]);
}

#[test]
fn rethrow_after_logging_leaves_outer_handler_with_same_message() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"seed\"); } catch (RuntimeException e) { System.out.println(\"log\"); throw e; } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["log", "seed"]);
}

#[test]
fn rethrow_from_illegal_argument_preserves_message() {
    let out = run_main(
        "try { try { throw new IllegalArgumentException(\"arg\"); } catch (IllegalArgumentException e) { throw e; } } catch (IllegalArgumentException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["arg"]);
}

#[test]
fn catch_and_wrap_changes_outer_exception_message() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"cause\"); } catch (RuntimeException e) { throw new RuntimeException(\"wrapped\"); } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["wrapped"]);
}

#[test]
fn catch_and_wrap_with_concatenated_cause_message() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"root\"); } catch (RuntimeException e) { throw new RuntimeException(\"wrap:\" + e.getMessage()); } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["wrap:root"]);
}

#[test]
fn catch_and_wrap_io_exception_into_runtime_exception() {
    let out = run_main(
        "try { try { throw new java.io.IOException(\"io\"); } catch (java.io.IOException e) { throw new RuntimeException(\"failed:\" + e.getMessage()); } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["failed:io"]);
}

#[test]
fn multi_catch_runtime_or_illegal_argument_matches_runtime() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException | IllegalArgumentException e) { System.out.println(\"union-runtime\"); }",
    );
    assert_eq!(out, vec!["union-runtime"]);
}

#[test]
fn multi_catch_runtime_or_illegal_argument_matches_illegal() {
    let out = run_main(
        "try { throw new IllegalArgumentException(); } catch (RuntimeException | IllegalArgumentException e) { System.out.println(\"union-illegal\"); }",
    );
    assert_eq!(out, vec!["union-illegal"]);
}

#[test]
fn multi_catch_null_or_index_matches_null_pointer() {
    let out = run_main(
        "try { throw new NullPointerException(); } catch (NullPointerException | IndexOutOfBoundsException e) { System.out.println(\"union-npe\"); }",
    );
    assert_eq!(out, vec!["union-npe"]);
}

#[test]
fn multi_catch_null_or_index_matches_index_out_of_bounds() {
    let out = run_main(
        "try { throw new IndexOutOfBoundsException(); } catch (NullPointerException | IndexOutOfBoundsException e) { System.out.println(\"union-index\"); }",
    );
    assert_eq!(out, vec!["union-index"]);
}

#[test]
fn multi_catch_io_or_runtime_matches_io_only() {
    let out = run_main(
        "try { throw new java.io.IOException(); } catch (java.io.IOException | RuntimeException e) { System.out.println(\"union-io\"); }",
    );
    assert_eq!(out, vec!["union-io"]);
}

#[test]
fn multi_catch_with_exception_supertype_covers_custom_error() {
    let types = r#"
        static class ParseError extends Exception {
            ParseError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new ParseError(\"bad\"); } catch (ParseError | RuntimeException e) { System.out.println(\"parsed\"); }",
        types,
    );
    assert_eq!(out, vec!["parsed"]);
}

#[test]
fn init_cause_links_previous_exception_message() {
    let out = run_main(
        "try { RuntimeException root = new RuntimeException(\"root\"); RuntimeException top = new RuntimeException(\"top\"); top.initCause(root); throw top; } catch (RuntimeException e) { System.out.println(e.getCause().getMessage()); }",
    );
    assert_eq!(out, vec!["root"]);
}

#[test]
fn init_cause_returns_self_for_chaining() {
    let out = run_main(
        "try { RuntimeException root = new RuntimeException(\"r\"); RuntimeException top = new RuntimeException(\"t\"); top.initCause(root); throw top; } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["t"]);
}

#[test]
fn get_cause_null_when_no_cause_was_set() {
    let out = run_main(
        "try { throw new RuntimeException(\"solo\"); } catch (RuntimeException e) { System.out.println(e.getCause() == null); }",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn constructor_with_cause_argument_sets_get_cause() {
    let out = run_main(
        "try { RuntimeException root = new RuntimeException(\"inner\"); RuntimeException top = new RuntimeException(\"outer\", root); throw top; } catch (RuntimeException e) { System.out.println(e.getCause().getMessage()); }",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn constructor_with_cause_preserves_outer_message() {
    let out = run_main(
        "try { RuntimeException root = new RuntimeException(\"inner\"); RuntimeException top = new RuntimeException(\"outer\", root); throw top; } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["outer"]);
}

#[test]
fn init_cause_on_wrapped_io_exception_exposes_root() {
    let out = run_main(
        "try { java.io.IOException root = new java.io.IOException(\"socket\"); RuntimeException top = new RuntimeException(\"network\"); top.initCause(root); throw top; } catch (RuntimeException e) { System.out.println(e.getCause().getMessage()); }",
    );
    assert_eq!(out, vec!["socket"]);
}

#[test]
fn throw_skips_statements_after_throw_in_same_try() {
    let out = run_main(
        "try { System.out.println(\"before\"); throw new RuntimeException(\"stop\"); System.out.println(\"never\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["before", "stop"]);
}

#[test]
fn throw_inside_nested_try_reaches_outer_catch_when_inner_mismatches() {
    let out = run_main(
        "try { try { throw new IllegalArgumentException(\"x\"); } catch (NullPointerException e) { System.out.println(\"wrong\"); } } catch (IllegalArgumentException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn throw_custom_checked_exception_with_message() {
    let types = r#"
        static class DomainError extends Exception {
            DomainError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new DomainError(\"domain\"); } catch (DomainError e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["domain"]);
}

#[test]
fn rethrow_custom_exception_preserves_message() {
    let types = r#"
        static class AppError extends Exception {
            AppError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { try { throw new AppError(\"app\"); } catch (AppError e) { throw e; } } catch (AppError e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["app"]);
}

#[test]
fn catch_and_wrap_custom_exception_changes_message() {
    let types = r#"
        static class LowError extends Exception {
            LowError(String msg) { super(msg); }
        }
        static class HighError extends Exception {
            HighError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { try { throw new LowError(\"low\"); } catch (LowError e) { throw new HighError(\"high\"); } } catch (HighError e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["high"]);
}

#[test]
fn multi_catch_three_type_union_matches_middle_type() {
    let out = run_main(
        "try { throw new IllegalArgumentException(); } catch (RuntimeException | IllegalArgumentException | NullPointerException e) { System.out.println(\"triple\"); }",
    );
    assert_eq!(out, vec!["triple"]);
}

#[test]
fn throw_in_loop_caught_each_iteration_independently() {
    let out = run_main(
        "for (int i = 0; i < 3; i++) { try { if (i == 1) throw new RuntimeException(\"hit\"); System.out.println(i); } catch (RuntimeException e) { System.out.println(e.getMessage()); } }",
    );
    assert_eq!(out, vec!["0", "hit", "2"]);
}

#[test]
fn rethrow_sets_recovery_flag_before_outer_catch() {
    let out = run_main(
        "int recovered = 0; try { try { throw new RuntimeException(); } catch (RuntimeException e) { recovered = 1; throw e; } } catch (RuntimeException e) { System.out.println(recovered); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn catch_and_wrap_then_get_cause_from_constructor_form() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"leaf\"); } catch (RuntimeException e) { throw new RuntimeException(\"branch\", e); } } catch (RuntimeException e) { System.out.println(e.getCause().getMessage()); }",
    );
    assert_eq!(out, vec!["leaf"]);
}

#[test]
fn throw_new_exception_after_assigning_catch_variable_name() {
    let out = run_main(
        "try { throw new RuntimeException(\"msg\"); } catch (RuntimeException ex) { System.out.println(ex.getMessage()); }",
    );
    assert_eq!(out, vec!["msg"]);
}

#[test]
fn multi_catch_union_does_not_match_unrelated_type() {
    let out = run_main(
        "try { try { throw new java.io.IOException(\"io\"); } catch (RuntimeException | IllegalArgumentException e) { System.out.println(\"wrong\"); } } catch (java.io.IOException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["io"]);
}

#[test]
fn throw_declared_runtime_inside_finally_after_catch() {
    let out = run_main(
        "try { throw new RuntimeException(\"try\"); } catch (RuntimeException e) { System.out.println(\"catch\"); } finally { System.out.println(\"finally\"); }",
    );
    assert_eq!(out, vec!["catch", "finally"]);
}

#[test]
fn init_cause_then_rethrow_preserves_cause_for_outer_handler() {
    let out = run_main(
        "try { try { RuntimeException r = new RuntimeException(\"base\"); RuntimeException w = new RuntimeException(\"wrap\"); w.initCause(r); throw w; } catch (RuntimeException e) { throw e; } } catch (RuntimeException e) { System.out.println(e.getCause().getMessage()); }",
    );
    assert_eq!(out, vec!["base"]);
}

#[test]
fn throw_new_after_successful_try_body_in_sequence() {
    let out = run_main(
        "try { System.out.println(\"ok\"); } catch (RuntimeException e) { System.out.println(\"no\"); } try { throw new RuntimeException(\"later\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["ok", "later"]);
}

#[test]
fn catch_and_wrap_null_pointer_into_runtime_with_message() {
    let out = run_main(
        "try { try { throw new NullPointerException(\"n\"); } catch (NullPointerException e) { throw new RuntimeException(\"r:\" + e.getMessage()); } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["r:n"]);
}

#[test]
fn get_cause_on_io_exception_constructor_chain() {
    let out = run_main(
        "try { java.io.IOException root = new java.io.IOException(\"disk\"); java.io.IOException top = new java.io.IOException(\"read\", root); throw top; } catch (java.io.IOException e) { System.out.println(e.getCause().getMessage()); }",
    );
    assert_eq!(out, vec!["disk"]);
}
