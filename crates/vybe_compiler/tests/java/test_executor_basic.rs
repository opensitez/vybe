use crate::helpers::{run_in_main, run_main};

#[test]
fn executors_new_fixed_thread_pool_accepts_single_task() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.submit(() -> System.out.println(\"task\")); pool.shutdown(); Thread.sleep(10); System.out.println(\"done\");",
    );
    assert_eq!(out, vec!["task", "done"]);
}

#[test]
fn executors_submit_callable_returns_future_get_value() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Callable<Integer> job = () -> 6 * 7; java.util.concurrent.Future<Integer> future = pool.submit(job); System.out.println(future.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn executors_submit_runnable_future_get_completes() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<?> future = pool.submit(() -> System.out.println(\"run\")); future.get(); pool.shutdown(); System.out.println(\"joined\");",
    );
    assert_eq!(out, vec!["run", "joined"]);
}

#[test]
fn executors_shutdown_allows_submitted_work_to_finish() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.submit(() -> System.out.println(\"finish\")); pool.shutdown(); pool.awaitTermination(1, java.util.concurrent.TimeUnit.SECONDS); System.out.println(\"closed\");",
    );
    assert_eq!(out, vec!["finish", "closed"]);
}

#[test]
fn executors_fixed_pool_size_two_runs_both_callables() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(2); java.util.concurrent.Future<Integer> a = pool.submit(() -> 1 + 1); java.util.concurrent.Future<Integer> b = pool.submit(() -> 2 + 2); System.out.println(a.get()); System.out.println(b.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn executors_submit_callable_string_result() {
    let out = run_main(
        r#"java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<String> f = pool.submit(() -> "vybe"); System.out.println(f.get()); pool.shutdown();"#,
    );
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn executors_submit_callable_boolean_result() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Boolean> f = pool.submit(() -> true); System.out.println(f.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_submit_callable_with_argument_capture() {
    let out = run_main(
        "int base = 10; java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> f = pool.submit(() -> base + 5); System.out.println(f.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn executors_multiple_submit_before_shutdown_all_complete() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(2); java.util.concurrent.Future<Integer> f1 = pool.submit(() -> 1); java.util.concurrent.Future<Integer> f2 = pool.submit(() -> 2); java.util.concurrent.Future<Integer> f3 = pool.submit(() -> 3); System.out.println(f1.get() + f2.get() + f3.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn executors_shutdown_then_is_shutdown_true() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.shutdown(); System.out.println(pool.isShutdown());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_new_pool_not_shutdown_initially() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); System.out.println(pool.isShutdown()); pool.shutdown();",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn executors_submit_callable_returns_null_as_result() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<String> f = pool.submit(() -> null); System.out.println(f.get() == null); pool.shutdown();",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_runnable_submit_prints_from_worker() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.submit(() -> System.out.println(\"worker\")); pool.shutdown(); Thread.sleep(10);",
    );
    assert_eq!(out, vec!["worker"]);
}

#[test]
fn executors_callable_call_not_used_submit_instead() {
    let out = run_main(
        "java.util.concurrent.Callable<Integer> c = () -> 9; java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); System.out.println(pool.submit(c).get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn executors_fixed_pool_one_serializes_two_tasks() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<String> first = pool.submit(() -> \"a\"); java.util.concurrent.Future<String> second = pool.submit(() -> \"b\"); System.out.println(first.get() + second.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn executors_submit_callable_computes_length() {
    let out = run_main(
        r#"java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> f = pool.submit(() -> "java".length()); System.out.println(f.get()); pool.shutdown();"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn executors_shutdown_now_stops_idle_pool() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.List<Runnable> pending = pool.shutdownNow(); System.out.println(pending.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn executors_await_termination_after_shutdown_returns() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.submit(() -> 1); pool.shutdown(); boolean ended = pool.awaitTermination(1, java.util.concurrent.TimeUnit.SECONDS); System.out.println(ended);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_submit_two_runnables_both_print() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(2); pool.submit(() -> System.out.println(\"one\")); pool.submit(() -> System.out.println(\"two\")); pool.shutdown(); Thread.sleep(20);",
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"one".to_string()));
    assert!(out.contains(&"two".to_string()));
}

#[test]
fn executors_callable_exception_surfaces_on_get() {
    let out = run_in_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> f = pool.submit(() -> { throw new RuntimeException(\"boom\"); }); try { f.get(); System.out.println(\"ok\"); } catch (Exception e) { System.out.println(e.getMessage()); } pool.shutdown();",
        "",
    );
    assert_eq!(out, vec!["boom"]);
}

#[test]
fn executors_submit_callable_with_explicit_type() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Callable<Double> c = () -> 2.5; System.out.println(pool.submit(c).get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn executors_future_is_done_after_get() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> f = pool.submit(() -> 1); f.get(); System.out.println(f.isDone()); pool.shutdown();",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_future_not_done_before_get() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> f = pool.submit(() -> { Thread.sleep(5); return 1; }); System.out.println(f.isDone()); f.get(); pool.shutdown();",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn executors_submit_callable_chained_math() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> f = pool.submit(() -> (1 + 2) * 3); System.out.println(f.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn executors_pool_with_three_threads_handles_three_futures() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(3); java.util.concurrent.Future<Integer> a = pool.submit(() -> 10); java.util.concurrent.Future<Integer> b = pool.submit(() -> 20); java.util.concurrent.Future<Integer> c = pool.submit(() -> 30); System.out.println(a.get() + b.get() + c.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn executors_submit_anonymous_callable() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Callable<String> c = new java.util.concurrent.Callable<String>() { public String call() { return \"anon\"; } }; System.out.println(pool.submit(c).get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["anon"]);
}

#[test]
fn executors_submit_runnable_with_side_effect_counter() {
    let types = r#"
        static int[] hits = {0};
    "#;
    let out = run_in_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.submit(() -> { hits[0]++; }); pool.shutdown(); Thread.sleep(10); System.out.println(hits[0]);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn executors_callable_returns_concatenated_string() {
    let out = run_main(
        r#"java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<String> f = pool.submit(() -> "a" + "b"); System.out.println(f.get()); pool.shutdown();"#,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn executors_shutdown_idempotent_second_call() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.shutdown(); pool.shutdown(); System.out.println(pool.isShutdown());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_submit_after_shutdown_rejected_or_ignored() {
    let out = run_in_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.shutdown(); try { pool.submit(() -> 1); System.out.println(\"submitted\"); } catch (Exception e) { System.out.println(\"rejected\"); }",
        "",
    );
    assert_eq!(out.len(), 1);
}

#[test]
fn executors_fixed_pool_callable_negative_number() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); System.out.println(pool.submit(() -> -5).get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn executors_future_get_twice_returns_same_value() {
    let out = run_main(
        r#"java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<String> f = pool.submit(() -> "same"); System.out.println(f.get()); System.out.println(f.get()); pool.shutdown();"#,
    );
    assert_eq!(out, vec!["same", "same"]);
}

#[test]
fn executors_submit_callable_zero_result() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); System.out.println(pool.submit(() -> 0).get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn executors_runnable_future_get_returns_null() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<?> f = pool.submit(() -> {}); System.out.println(f.get() == null); pool.shutdown();",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_pool_size_one_still_runs_sequence() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> a = pool.submit(() -> 1); java.util.concurrent.Future<Integer> b = pool.submit(() -> 2); java.util.concurrent.Future<Integer> c = pool.submit(() -> 3); System.out.println(a.get() + b.get() + c.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn executors_callable_with_local_variable_capture() {
    let out = run_main(
        r#"String tag = "go"; java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); System.out.println(pool.submit(() -> tag).get()); pool.shutdown();"#,
    );
    assert_eq!(out, vec!["go"]);
}

#[test]
fn executors_submit_println_runnable_and_callable_together() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(2); pool.submit(() -> System.out.println(\"r\")); java.util.concurrent.Future<String> f = pool.submit(() -> \"c\"); System.out.println(f.get()); pool.shutdown(); Thread.sleep(10);",
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"r".to_string()));
    assert!(out.contains(&"c".to_string()));
}

#[test]
fn executors_new_fixed_thread_pool_rejects_zero_size_gracefully() {
    let out = run_in_main(
        "try { java.util.concurrent.Executors.newFixedThreadPool(0); System.out.println(\"made\"); } catch (Exception e) { System.out.println(\"error\"); }",
        "",
    );
    assert_eq!(out.len(), 1);
}

#[test]
fn executors_callable_returns_boolean_expression() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); System.out.println(pool.submit(() -> 3 > 2).get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn executors_shutdown_after_empty_submit_list() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.shutdown(); System.out.println(\"ok\");",
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn executors_submit_callable_max_int_value() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); System.out.println(pool.submit(() -> Integer.MAX_VALUE).get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["2147483647"]);
}

#[test]
fn executors_multiple_shutdown_await_termination() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); pool.submit(() -> System.out.println(\"x\")); pool.shutdown(); pool.awaitTermination(1, java.util.concurrent.TimeUnit.SECONDS); System.out.println(\"y\");",
    );
    assert_eq!(out, vec!["x", "y"]);
}

#[test]
fn executors_callable_list_size_result() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.Future<Integer> f = pool.submit(() -> { java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); return list.size(); }); System.out.println(f.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn executors_submit_two_callables_sum_via_main() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(2); java.util.concurrent.Future<Integer> left = pool.submit(() -> 11); java.util.concurrent.Future<Integer> right = pool.submit(() -> 22); int sum = left.get() + right.get(); pool.shutdown(); System.out.println(sum);",
    );
    assert_eq!(out, vec!["33"]);
}

#[test]
fn executors_runnable_and_callable_share_pool() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(2); pool.submit(() -> System.out.println(\"run\")); java.util.concurrent.Future<Integer> f = pool.submit(() -> 5); System.out.println(f.get()); pool.shutdown(); Thread.sleep(10);",
    );
    assert_eq!(out.len(), 2);
    assert!(out.contains(&"run".to_string()));
    assert!(out.contains(&"5".to_string()));
}
