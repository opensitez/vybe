/// Suppressed exceptions: addSuppressed / getSuppressed.
use crate::helpers::{run_in_main, run_main};

#[test]
fn try_with_resources_suppresses_close_exception_when_body_throws() {
    let out = run_in_main(
        "try (NoisyCloser c = new NoisyCloser()) { throw new RuntimeException(\"body\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); System.out.println(e.getSuppressed().length); }",
        r#"
        static class NoisyCloser implements AutoCloseable {
            public void close() { throw new IllegalStateException("close"); }
        }
        "#,
    );
    assert_eq!(out, vec!["body", "1"]);
}

#[test]
fn catch_block_can_attach_suppressed_exception_manually() {
    let out = run_main(
        "try { throw new RuntimeException(\"primary\"); } catch (RuntimeException e) { e.addSuppressed(new IllegalArgumentException(\"extra\")); System.out.println(e.getSuppressed()[0].getClass().getSimpleName()); }",
    );
    assert_eq!(out, vec!["IllegalArgumentException"]);
}

#[test]
fn suppressed_array_empty_when_none_added() {
    let out = run_main(
        "try { throw new RuntimeException(\"only\"); } catch (RuntimeException e) { System.out.println(e.getSuppressed().length); }",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn multiple_suppressed_exceptions_preserve_order() {
    let out = run_main(
        "try { throw new RuntimeException(\"main\"); } catch (RuntimeException e) { e.addSuppressed(new IllegalStateException(\"first\")); e.addSuppressed(new IllegalArgumentException(\"second\")); System.out.println(e.getSuppressed()[0].getMessage()); System.out.println(e.getSuppressed()[1].getMessage()); }",
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn suppressed_exception_message_readable_after_add() {
    let out = run_main(
        r#"try { throw new RuntimeException("outer"); } catch (RuntimeException e) { e.addSuppressed(new RuntimeException("inner")); System.out.println(e.getSuppressed()[0].getMessage()); }"#,
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn rethrown_exception_retains_suppressed_from_cause_chain() {
    let out = run_main(
        "try { try { throw new RuntimeException(\"root\"); } catch (RuntimeException inner) { inner.addSuppressed(new IllegalStateException(\"tag\")); throw new RuntimeException(\"wrap\", inner); } } catch (RuntimeException outer) { System.out.println(outer.getCause().getSuppressed().length); }",
    );
    assert_eq!(out, vec!["1"]);
}
