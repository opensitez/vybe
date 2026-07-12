use crate::helpers::{run_in_main, run_main};

#[test]
fn runtime_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (Exception e) { System.out.println(\"runtime-via-exception\"); }",
    );
    assert_eq!(out, vec!["runtime-via-exception"]);
}

#[test]
fn runtime_exception_message_survives_exception_catch() {
    let out = run_main(
        "try { throw new RuntimeException(\"disk full\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["disk full"]);
}

#[test]
fn illegal_argument_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new IllegalArgumentException(); } catch (Exception e) { System.out.println(\"illegal-via-exception\"); }",
    );
    assert_eq!(out, vec!["illegal-via-exception"]);
}

#[test]
fn illegal_argument_exception_message_via_exception_handler() {
    let out = run_main(
        "try { throw new IllegalArgumentException(\"negative width\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["negative width"]);
}

#[test]
fn null_pointer_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new NullPointerException(); } catch (Exception e) { System.out.println(\"npe-via-exception\"); }",
    );
    assert_eq!(out, vec!["npe-via-exception"]);
}

#[test]
fn null_pointer_exception_message_via_exception_handler() {
    let out = run_main(
        "try { throw new NullPointerException(\"ref was null\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["ref was null"]);
}

#[test]
fn index_out_of_bounds_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new IndexOutOfBoundsException(); } catch (Exception e) { System.out.println(\"ioob-via-exception\"); }",
    );
    assert_eq!(out, vec!["ioob-via-exception"]);
}

#[test]
fn index_out_of_bounds_exception_message_via_exception_handler() {
    let out = run_main(
        "try { throw new IndexOutOfBoundsException(\"index 9\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["index 9"]);
}

#[test]
fn io_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new java.io.IOException(); } catch (Exception e) { System.out.println(\"io-via-exception\"); }",
    );
    assert_eq!(out, vec!["io-via-exception"]);
}

#[test]
fn io_exception_message_via_exception_handler() {
    let out = run_main(
        "try { throw new java.io.IOException(\"read failed\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["read failed"]);
}

#[test]
fn runtime_exception_caught_by_runtime_exception_type() {
    let out = run_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"runtime-direct\"); }",
    );
    assert_eq!(out, vec!["runtime-direct"]);
}

#[test]
fn illegal_argument_caught_by_runtime_exception_supertype() {
    let out = run_main(
        "try { throw new IllegalArgumentException(); } catch (RuntimeException e) { System.out.println(\"illegal-as-runtime\"); }",
    );
    assert_eq!(out, vec!["illegal-as-runtime"]);
}

#[test]
fn null_pointer_caught_by_runtime_exception_supertype() {
    let out = run_main(
        "try { throw new NullPointerException(); } catch (RuntimeException e) { System.out.println(\"npe-as-runtime\"); }",
    );
    assert_eq!(out, vec!["npe-as-runtime"]);
}

#[test]
fn index_out_of_bounds_caught_by_runtime_exception_supertype() {
    let out = run_main(
        "try { throw new IndexOutOfBoundsException(); } catch (RuntimeException e) { System.out.println(\"ioob-as-runtime\"); }",
    );
    assert_eq!(out, vec!["ioob-as-runtime"]);
}

#[test]
fn io_exception_not_caught_by_runtime_exception_handler() {
    let out = run_main(
        "try { try { throw new java.io.IOException(\"eof\"); } catch (RuntimeException e) { System.out.println(\"wrong\"); } } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["eof"]);
}

#[test]
fn array_index_out_of_bounds_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new ArrayIndexOutOfBoundsException(\"slot 3\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["slot 3"]);
}

#[test]
fn string_index_out_of_bounds_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new StringIndexOutOfBoundsException(\"char 12\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["char 12"]);
}

#[test]
fn number_format_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new NumberFormatException(\"not a number\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["not a number"]);
}

#[test]
fn unsupported_operation_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new UnsupportedOperationException(); } catch (Exception e) { System.out.println(\"unsupported\"); }",
    );
    assert_eq!(out, vec!["unsupported"]);
}

#[test]
fn arithmetic_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new ArithmeticException(\"divide by zero\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["divide by zero"]);
}

#[test]
fn class_cast_exception_caught_by_exception_supertype() {
    let out = run_main(
        "try { throw new ClassCastException(\"bad cast\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["bad cast"]);
}

#[test]
fn exception_handler_runs_after_runtime_exception_throw() {
    let out = run_main(
        "try { System.out.println(\"before\"); throw new RuntimeException(); } catch (Exception e) { System.out.println(\"after\"); }",
    );
    assert_eq!(out, vec!["before", "after"]);
}

#[test]
fn exception_handler_preserves_local_state_set_before_throw() {
    let out = run_main(
        "int code = 0; try { code = 42; throw new IllegalArgumentException(); } catch (Exception e) { System.out.println(code); }",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn nested_exception_catch_inner_runtime_outer_exception() {
    let out = run_main(
        "try { try { throw new NullPointerException(\"inner\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); } } catch (Exception e) { System.out.println(\"outer\"); }",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn nested_exception_catch_inner_io_outer_exception() {
    let out = run_main(
        "try { try { throw new java.io.IOException(\"nested\"); } catch (RuntimeException e) { System.out.println(\"wrong\"); } } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["nested"]);
}

#[test]
fn custom_runtime_subclass_caught_by_exception() {
    let types = r#"
        static class ConfigError extends RuntimeException {
            ConfigError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new ConfigError(\"bad config\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["bad config"]);
}

#[test]
fn custom_checked_subclass_caught_by_exception() {
    let types = r#"
        static class ParseError extends Exception {
            ParseError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new ParseError(\"syntax\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["syntax"]);
}

#[test]
fn custom_io_subclass_caught_by_exception() {
    let types = r#"
        static class ReadError extends java.io.IOException {
            ReadError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new ReadError(\"short read\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["short read"]);
}

#[test]
fn exception_catch_all_handles_null_pointer_in_loop() {
    let out = run_main(
        "for (int i = 0; i < 2; i++) { try { if (i == 1) throw new NullPointerException(\"loop\"); System.out.println(i); } catch (Exception e) { System.out.println(e.getMessage()); } }",
    );
    assert_eq!(out, vec!["0", "loop"]);
}

#[test]
fn exception_catch_all_handles_index_error_in_loop() {
    let out = run_main(
        "for (int i = 0; i < 2; i++) { try { if (i == 1) throw new IndexOutOfBoundsException(\"idx\"); System.out.println(i); } catch (Exception e) { System.out.println(e.getMessage()); } }",
    );
    assert_eq!(out, vec!["0", "idx"]);
}

#[test]
fn exception_catch_all_handles_io_error_in_loop() {
    let out = run_main(
        "for (int i = 0; i < 2; i++) { try { if (i == 1) throw new java.io.IOException(\"io\"); System.out.println(i); } catch (Exception e) { System.out.println(e.getMessage()); } }",
    );
    assert_eq!(out, vec!["0", "io"]);
}

#[test]
fn runtime_exception_empty_message_has_zero_length_in_handler() {
    let out = run_main(
        "try { throw new RuntimeException(\"\"); } catch (Exception e) { System.out.println(e.getMessage().length()); }",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn illegal_argument_concatenated_message_reaches_exception_handler() {
    let out = run_main(
        "try { throw new IllegalArgumentException(\"bad:\" + 7); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["bad:7"]);
}

#[test]
fn null_pointer_from_explicit_throw_not_from_null_deref() {
    let out = run_main(
        "try { throw new NullPointerException(\"explicit\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["explicit"]);
}

#[test]
fn index_out_of_bounds_from_array_access_pattern() {
    let out = run_main(
        "int[] data = {1, 2}; try { throw new ArrayIndexOutOfBoundsException(\"\" + 5); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn io_exception_subclass_file_not_found_caught_by_exception() {
    let out = run_main(
        "try { throw new java.io.FileNotFoundException(\"missing.txt\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["missing.txt"]);
}

#[test]
fn exception_supertype_catches_after_specific_runtime_handler_skips() {
    let out = run_main(
        "try { throw new java.io.IOException(\"stream\"); } catch (RuntimeException e) { System.out.println(\"skip\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["stream"]);
}

#[test]
fn two_sequential_exception_handlers_are_independent() {
    let out = run_main(
        "try { throw new RuntimeException(\"a\"); } catch (Exception e) { System.out.println(e.getMessage()); } try { throw new IllegalArgumentException(\"b\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["a", "b"]);
}

#[test]
fn exception_handler_continues_after_catch_block() {
    let out = run_main(
        "try { throw new IndexOutOfBoundsException(); } catch (Exception e) { System.out.println(\"handled\"); } System.out.println(\"next\");",
    );
    assert_eq!(out, vec!["handled", "next"]);
}

#[test]
fn runtime_and_io_exceptions_both_route_to_exception_in_sequence() {
    let out = run_main(
        "try { throw new RuntimeException(\"r\"); } catch (Exception e) { System.out.println(e.getMessage()); } try { throw new java.io.IOException(\"i\"); } catch (Exception e) { System.out.println(e.getMessage()); }",
    );
    assert_eq!(out, vec!["r", "i"]);
}
