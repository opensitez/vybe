use crate::helpers::{run_in_main, run_main};

#[test]
fn try_body_runs_when_no_exception_thrown() {
    let out = run_main("try { System.out.println(\"ok\"); } catch (RuntimeException e) { System.out.println(\"no\"); }");
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn catch_handles_thrown_runtime_exception() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"caught\"); }",
    );
    assert_eq!(out, vec!["caught"]);
}

#[test]
fn finally_runs_after_successful_try_without_catch() {
    let out = run_main(
        "try { System.out.println(\"try\"); } finally { System.out.println(\"finally\"); }",
    );
    assert_eq!(out, vec!["try", "finally"]);
}

#[test]
fn finally_runs_after_catch_handles_exception() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"catch\"); } finally { System.out.println(\"finally\"); }",
    );
    assert_eq!(out, vec!["catch", "finally"]);
}

#[test]
fn finally_skipped_when_no_exception_and_no_finally_clause() {
    let out = run_main("try { System.out.println(\"only\"); } catch (RuntimeException e) { System.out.println(\"skip\"); }");
    assert_eq!(out, vec!["only"]);
}

#[test]
fn catch_supertype_handles_illegal_argument_subclass() {
    let out = run_main(
        "try { throw new IllegalArgumentException(); } catch (Exception e) { System.out.println(\"handled\"); }",
    );
    assert_eq!(out, vec!["handled"]);
}

#[test]
fn exception_get_message_returns_constructor_text() {
    let out = run_main(
        "try { throw new RuntimeException(\"disk full\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["disk full"]);
}

#[test]
fn exception_with_empty_string_message_has_zero_length() {
    let out = run_main(
        "try { throw new RuntimeException(\"\"); } catch (RuntimeException e) { System.out.println(e.getMessage().length()); }",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn first_matching_catch_arm_handles_runtime_exception() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"runtime\"); } catch (Exception e) { System.out.println(\"general\"); }",
    );
    assert_eq!(out, vec!["runtime"]);
}

#[test]
fn second_catch_arm_handles_when_first_type_mismatches() {
    let out = run_main(
        "try { throw new IllegalArgumentException(); } catch (RuntimeException e) { System.out.println(\"runtime\"); } catch (IllegalArgumentException e) { System.out.println(\"illegal\"); }",
    );
    assert_eq!(out, vec!["illegal"]);
}

#[test]
fn catch_order_specific_type_wins_over_exception_supertype() {
    let out = run_main(
        "try { throw new IllegalArgumentException(\"bad\"); } catch (IllegalArgumentException e) { System.out.println(e.getMessage()); } catch (Exception e) { System.out.println(\"fallback\"); }",
    );
    assert_eq!(out, vec!["bad"]);
}

#[test]
fn multi_catch_union_matches_runtime_exception_arm() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException | IllegalArgumentException e) { System.out.println(\"union\"); }",
    );
    assert_eq!(out, vec!["union"]);
}

#[test]
fn multi_catch_union_matches_illegal_argument_variant() {
    let out = run_main(
        "try { throw new IllegalArgumentException(); } catch (RuntimeException | IllegalArgumentException e) { System.out.println(\"union\"); }",
    );
    assert_eq!(out, vec!["union"]);
}

#[test]
fn catch_all_exception_arm_handles_any_throwable() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (Exception e) { System.out.println(\"any\"); }",
    );
    assert_eq!(out, vec!["any"]);
}

#[test]
fn rethrow_from_catch_preserves_message_for_outer_handler() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"inner\"); } catch (RuntimeException e) { throw e; } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn rethrow_wrapped_in_new_exception_changes_outer_message() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"cause\"); } catch (RuntimeException e) { throw new RuntimeException(\"wrapped\"); } } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["wrapped"]);
}

#[test]
fn nested_try_inner_catch_handles_without_reaching_outer() {
    let out = run_main(
        "try { try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"inner\"); } } catch (RuntimeException e) { System.out.println(\"outer\"); }",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn nested_try_outer_catch_handles_uncaught_inner_throw() {
    let out = run_main(
        "try { try { throw new RuntimeException(); } catch (IllegalArgumentException e) { System.out.println(\"wrong\"); } } catch (RuntimeException e) { System.out.println(\"outer\"); }",
    );
    assert_eq!(out, vec!["outer"]);
}

#[test]
fn catch_variable_is_visible_inside_catch_body() {
    let out = run_main(
        "try { throw new RuntimeException(\"msg\"); } catch (RuntimeException ex) { System.out.println(ex.getMessage()); }",
    );
    assert_eq!(out, vec!["msg"]);
}

#[test]
fn try_assigns_local_before_throw_and_catch_reads_state() {
    let out = run_main(
        "int flag = 0; try { flag = 1; throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(flag); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn execution_continues_after_catch_block_completes() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"handled\"); } System.out.println(\"after\");",
    );
    assert_eq!(out, vec!["handled", "after"]);
}

#[test]
fn finally_runs_before_code_after_try_on_success_path() {
    let out = run_main(
        "try { System.out.println(\"try\"); } finally { System.out.println(\"finally\"); } System.out.println(\"after\");",
    );
    assert_eq!(out, vec!["try", "finally", "after"]);
}

#[test]
fn finally_runs_before_code_after_try_on_catch_path() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"catch\"); } finally { System.out.println(\"finally\"); } System.out.println(\"after\");",
    );
    assert_eq!(out, vec!["catch", "finally", "after"]);
}

#[test]
fn try_with_resources_syntax_executes_body() {
    let types = r#"
        static class QuietCloser implements AutoCloseable {
            public void close() { System.out.println("closed"); }
        }
    "#;
    let out = run_in_main(
        "try (QuietCloser resource = new QuietCloser()) { System.out.println(\"body\"); }",
        types,
    );
    assert_eq!(out, vec!["body"]);
}

#[test]
fn try_with_resources_with_catch_handles_exception_in_body() {
    let types = r#"
        static class QuietCloser implements AutoCloseable {
            public void close() { }
        }
    "#;
    let out = run_in_main(
        "try (QuietCloser resource = new QuietCloser()) { throw new RuntimeException(\"io\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["io"]);
}

#[test]
fn custom_checked_exception_message_survives_to_catch() {
    let types = r#"
        static class ConfigError extends Exception {
            ConfigError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new ConfigError(\"missing key\"); } catch (ConfigError e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["missing key"]);
}

#[test]
fn custom_exception_caught_by_declared_superclass() {
    let types = r#"
        static class AppError extends Exception {
            AppError(String msg) { super(msg); }
        }
        static class NotFound extends AppError {
            NotFound(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new NotFound(\"404\"); } catch (AppError e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["404"]);
}

#[test]
fn custom_exception_caught_by_exception_supertype() {
    let types = r#"
        static class DomainError extends Exception {
            DomainError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new DomainError(\"fail\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["fail"]);
}

#[test]
fn throw_after_catch_setup_in_same_try_block() {
    let out = run_main(
        "try { System.out.println(\"before\"); throw new RuntimeException(\"stop\"); System.out.println(\"never\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["before", "stop"]);
}

#[test]
fn concatenated_message_in_thrown_runtime_exception() {
    let out = run_main(
        "try { throw new RuntimeException(\"err:\" + 42); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["err:42"]);
}

#[test]
fn two_sequential_try_catch_blocks_are_independent() {
    let out = run_main(
        "try { throw new RuntimeException(\"a\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); } try { throw new RuntimeException(\"b\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn inner_finally_runs_before_outer_catch_handles() {
    let out = run_main(
        "try { try { throw new RuntimeException(); } finally { System.out.println(\"inner-finally\"); } } catch (RuntimeException e) { System.out.println(\"outer-catch\"); }",
    );
    assert_eq!(out, vec!["inner-finally", "outer-catch"]);
}

#[test]
fn catch_runtime_then_nested_finally_in_same_try() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"catch\"); } finally { System.out.println(\"finally\"); }",
    );
    assert_eq!(out, vec!["catch", "finally"]);
}

#[test]
fn illegal_argument_message_propagates_to_handler() {
    let out = run_main(
        "try { throw new IllegalArgumentException(\"negative\"); } catch (IllegalArgumentException e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["negative"]);
}

#[test]
fn catch_logs_then_rethrows_for_outer_exception_handler() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"deep\"); } catch (RuntimeException e) { System.out.println(\"log\"); throw e; } } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["log", "deep"]);
}

#[test]
fn try_catch_inside_loop_handles_each_throw() {
    let out = run_main(
        "for (int i = 0; i < 3; i++) { try { if (i == 1) throw new RuntimeException(\"x\"); System.out.println(i); } catch (RuntimeException e) { System.out.println(\"hit\"); } }",
    );
    assert_eq!(out, vec!["0", "hit", "2"]);
}

#[test]
fn finally_executes_even_when_catch_body_runs() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"c\"); } finally { System.out.println(\"f\"); }",
    );
    assert_eq!(out, vec!["c", "f"]);
}

#[test]
fn multi_catch_with_exception_supertype_covers_custom_errors() {
    let types = r#"
        static class ParseError extends Exception {
            ParseError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new ParseError(\"syntax\"); } catch (ParseError | RuntimeException e) { System.out.println(\"parsed\"); }",
        types,
    );
    assert_eq!(out, vec!["parsed"]);
}

#[test]
fn rethrow_after_catch_assigns_recovery_flag() {
    let out = run_main(
        "int recovered = 0; try { try { throw new RuntimeException(); } catch (RuntimeException e) { recovered = 1; throw e; } } catch (RuntimeException e) { System.out.println(recovered); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn try_with_empty_catch_and_finally_on_success() {
    let out = run_main(
        "try { System.out.println(\"run\"); } catch (RuntimeException e) { } finally { System.out.println(\"done\"); }",
    );
    assert_eq!(out, vec!["run", "done"]);
}
