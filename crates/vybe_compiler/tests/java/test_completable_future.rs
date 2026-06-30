/// CompletableFuture composition and async supply.
use crate::helpers::run_main;

#[test]
fn completed_future_get_returns_value_immediately() {
    let out = run_main(
        r#"java.util.concurrent.CompletableFuture<String> f = java.util.concurrent.CompletableFuture.completedFuture("done"); System.out.println(f.get());"#,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn supply_async_computes_value_on_thread_pool() {
    let out = run_main(
        "java.util.concurrent.CompletableFuture<Integer> f = java.util.concurrent.CompletableFuture.supplyAsync(() -> 6 * 7); System.out.println(f.join());",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn then_apply_transforms_result() {
    let out = run_main(
        r#"java.util.concurrent.CompletableFuture<Integer> f = java.util.concurrent.CompletableFuture.completedFuture(3).thenApply(n -> n + 4); System.out.println(f.join());"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn then_combine_merges_two_futures() {
    let out = run_main(
        "java.util.concurrent.CompletableFuture<Integer> a = java.util.concurrent.CompletableFuture.completedFuture(10); java.util.concurrent.CompletableFuture<Integer> b = java.util.concurrent.CompletableFuture.completedFuture(5); System.out.println(a.thenCombine(b, (x, y) -> x - y).join());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn exceptionally_recovers_from_failed_stage() {
    let out = run_main(
        r#"java.util.concurrent.CompletableFuture<String> f = java.util.concurrent.CompletableFuture.<String>failedFuture(new RuntimeException("x")).exceptionally(ex -> "recovered"); System.out.println(f.join());"#,
    );
    assert_eq!(out, vec!["recovered"]);
}

#[test]
fn handle_provides_result_and_error_to_function() {
    let out = run_main(
        r#"java.util.concurrent.CompletableFuture<String> f = java.util.concurrent.CompletableFuture.completedFuture("ok").handle((val, err) -> val + "!"); System.out.println(f.join());"#,
    );
    assert_eq!(out, vec!["ok!"]);
}

#[test]
fn all_of_waits_for_both_stages() {
    let out = run_main(
        "java.util.concurrent.CompletableFuture<Void> all = java.util.concurrent.CompletableFuture.allOf(java.util.concurrent.CompletableFuture.completedFuture(1), java.util.concurrent.CompletableFuture.completedFuture(2)); all.join(); System.out.println(\"joined\");",
    );
    assert_eq!(out, vec!["joined"]);
}

#[test]
fn run_async_executes_runnable_without_return_value() {
    let out = run_main(
        "java.util.concurrent.CompletableFuture<Void> f = java.util.concurrent.CompletableFuture.runAsync(() -> System.out.println(\"ran\")); f.join();",
    );
    assert_eq!(out, vec!["ran"]);
}
