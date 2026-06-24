use crate::helpers::run_in_main;

#[test]
fn try_catch_handles_thrown_runtime_exception() {
    let out = run_in_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"caught\"); }",
        "",
    );
    assert_eq!(out, vec!["caught"]);
}

#[test]
fn finally_runs_after_catch() {
    let out = run_in_main(
        "try { throw new RuntimeException(); } catch (RuntimeException e) { System.out.println(\"catch\"); } finally { System.out.println(\"finally\"); }",
        "",
    );
    assert_eq!(out, vec!["catch", "finally"]);
}

#[test]
fn finally_runs_when_no_exception() {
    let out = run_in_main(
        "try { System.out.println(\"try\"); } catch (RuntimeException e) { System.out.println(\"catch\"); } finally { System.out.println(\"finally\"); }",
        "",
    );
    assert_eq!(out, vec!["try", "finally"]);
}

#[test]
fn catch_supertype_handles_subclass_exception() {
    let out = run_in_main(
        "try { throw new IllegalArgumentException(); } catch (Exception e) { System.out.println(\"handled\"); }",
        "",
    );
    assert_eq!(out, vec!["handled"]);
}

#[test]
fn thrown_exception_message_propagates() {
    let out = run_in_main(
        "try { throw new RuntimeException(\"boom\"); } catch (RuntimeException e) { System.out.println(e.getMessage()); }",
        "",
    );
    assert_eq!(out, vec!["boom"]);
}

#[test]
fn custom_checked_exception_extends_exception() {
    let types = r#"
        static class MyError extends Exception {
            MyError(String msg) { super(msg); }
        }
    "#;
    let out = run_in_main(
        "try { throw new MyError(\"fail\"); } catch (MyError e) { System.out.println(e.getMessage()); }",
        types,
    );
    assert_eq!(out, vec!["fail"]);
}
