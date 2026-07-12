/// java.util.concurrent.FutureTask — Runnable/Callable future adapter.
use crate::helpers::{run_in_main, run_main};

#[test]
fn future_task_callable_constructor_get_returns_result() {
    let out = run_main(
        "java.util.concurrent.Callable<Integer> job = () -> 6 * 7; java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(job); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn future_task_runnable_constructor_get_returns_null() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Void> task = new java.util.concurrent.FutureTask<Void>(() -> System.out.println(\"run\"), null); task.run(); System.out.println(task.get() == null);",
    );
    assert_eq!(out, vec!["run", "true"]);
}

#[test]
fn future_task_runnable_with_result_get_returns_value() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> {}, "done"); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn future_task_is_done_false_before_run() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); System.out.println(task.isDone());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn future_task_is_done_true_after_run() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); task.run(); System.out.println(task.isDone());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn future_task_is_cancelled_false_initially() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); System.out.println(task.isCancelled());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn future_task_cancel_before_run_prevents_execution() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "never"); task.cancel(false); System.out.println(task.isCancelled()); System.out.println(task.isDone());"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn future_task_cancel_after_run_has_no_effect_on_result() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 5); task.run(); task.cancel(false); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn future_task_run_executes_callable_body() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "executed"); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["executed"]);
}

#[test]
fn future_task_get_twice_returns_same_value() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "same"); task.run(); System.out.println(task.get()); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["same", "same"]);
}

#[test]
fn future_task_submit_to_executor_via_future_task() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 10); pool.submit(task); System.out.println(task.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn future_task_run_and_reset_allows_second_execution() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); task.run(); System.out.println(task.get()); task.runAndReset(); System.out.println(task.isDone());",
    );
    assert_eq!(out, vec!["1", "false"]);
}

#[test]
fn future_task_callable_exception_surfaces_on_get() {
    let out = run_in_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> { throw new RuntimeException(\"fail\"); }); task.run(); try { task.get(); System.out.println(\"ok\"); } catch (Exception e) { System.out.println(e.getCause().getMessage()); }",
        "",
    );
    assert_eq!(out, vec!["fail"]);
}

#[test]
fn future_task_cancel_may_interrupt_if_running_with_true() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> { Thread.sleep(50); return 1; }); Thread t = new Thread(task); t.start(); task.cancel(true); System.out.println(task.isCancelled());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn future_task_runnable_side_effect_before_result() {
    let types = r#"
        static int hits = 0;
    "#;
    let out = run_in_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> { hits++; }, 9); task.run(); System.out.println(hits); System.out.println(task.get());",
        types,
    );
    assert_eq!(out, vec!["1", "9"]);
}

#[test]
fn future_task_callable_returns_boolean() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Boolean> task = new java.util.concurrent.FutureTask<Boolean>(() -> 3 > 2); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn future_task_callable_returns_null() {
    let out = run_main(
        "java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> null); task.run(); System.out.println(task.get() == null);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn future_task_is_done_after_cancel_before_run() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); task.cancel(false); System.out.println(task.isDone());",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn future_task_get_after_cancel_throws_cancellation_exception() {
    let out = run_in_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); task.cancel(false); try { task.get(); System.out.println(\"ok\"); } catch (java.util.concurrent.CancellationException e) { System.out.println(\"cancelled\"); }",
        "",
    );
    assert_eq!(out, vec!["cancelled"]);
}

#[test]
fn future_task_run_on_new_thread_via_start_as_async() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 99); new Thread(task).start(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn future_task_callable_string_concatenation() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "a" + "b"); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn future_task_callable_with_local_capture() {
    let out = run_main(
        "int base = 10; java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> base + 5); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn future_task_multiple_run_only_first_counts() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "once"); task.run(); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["once"]);
}

#[test]
fn future_task_executor_submit_future_task_runnable() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> System.out.println(\"exec\"), \"ok\"); pool.execute(task); System.out.println(task.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["exec", "ok"]);
}

#[test]
fn future_task_callable_list_size_computation() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> { java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); list.add(1); list.add(2); return list.size(); }); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn future_task_is_cancelled_false_after_successful_run() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); task.run(); System.out.println(task.isCancelled());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn future_task_get_with_timeout_returns_result() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 7); task.run(); System.out.println(task.get(1, java.util.concurrent.TimeUnit.SECONDS));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn future_task_callable_negative_number_result() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> -5); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn future_task_callable_zero_result() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 0); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn future_task_callable_max_int_value() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> Integer.MAX_VALUE); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["2147483647"]);
}

#[test]
fn future_task_anonymous_callable_via_future_task() {
    let out = run_main(
        r#"java.util.concurrent.Callable<String> c = new java.util.concurrent.Callable<String>() { public String call() { return "anon"; } }; java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(c); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["anon"]);
}

#[test]
fn future_task_runnable_prints_during_run() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Void> task = new java.util.concurrent.FutureTask<Void>(() -> System.out.println(\"inside\"), null); task.run(); task.get();",
    );
    assert_eq!(out, vec!["inside"]);
}

#[test]
fn future_task_callable_double_result() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Double> task = new java.util.concurrent.FutureTask<Double>(() -> 2.5); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn future_task_cancel_false_does_not_interrupt_if_not_started() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 1); task.cancel(false); System.out.println(task.isCancelled()); System.out.println(task.isDone());",
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn future_task_run_and_reset_then_run_again() {
    let types = r#"
        static int runs = 0;
    "#;
    let out = run_in_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> { runs++; return runs; }); task.run(); System.out.println(task.get()); task.runAndReset(); task.run(); System.out.println(task.get());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn future_task_executor_runs_callable_future_task() {
    let out = run_main(
        r#"java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "pool"); pool.execute(task); System.out.println(task.get()); pool.shutdown();"#,
    );
    assert_eq!(out, vec!["pool"]);
}

#[test]
fn future_task_is_done_true_after_exception_in_run() {
    let out = run_in_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> { throw new RuntimeException(\"x\"); }); task.run(); System.out.println(task.isDone()); try { task.get(); } catch (Exception e) { System.out.println(\"err\"); }",
        "",
    );
    assert_eq!(out, vec!["true", "err"]);
}

#[test]
fn future_task_callable_boolean_false_result() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Boolean> task = new java.util.concurrent.FutureTask<Boolean>(() -> false); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn future_task_get_with_short_timeout_on_completed_task() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 3); task.run(); System.out.println(task.get(10, java.util.concurrent.TimeUnit.MILLISECONDS));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn future_task_runnable_result_overwrites_previous_get() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> {}, "first"); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn future_task_callable_math_expression() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> (1 + 2) * 3); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn future_task_thread_start_runs_asynchronously() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "async"); Thread t = new Thread(task); t.start(); t.join(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["async"]);
}

#[test]
fn future_task_cancelled_task_run_is_noop() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "run"); task.cancel(false); task.run(); System.out.println(task.isDone());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn future_task_callable_length_of_string() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> "java".length()); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn future_task_two_future_tasks_independent_results() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> a = new java.util.concurrent.FutureTask<Integer>(() -> 1); java.util.concurrent.FutureTask<Integer> b = new java.util.concurrent.FutureTask<Integer>(() -> 2); a.run(); b.run(); System.out.println(a.get() + b.get());",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn future_task_callable_explicit_generic_type() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Long> task = new java.util.concurrent.FutureTask<Long>(() -> 1000L); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["1000"]);
}

#[test]
fn future_task_runnable_void_result_null_on_get() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Void> task = new java.util.concurrent.FutureTask<Void>(() -> {}, null); task.run(); System.out.println(task.get() == null);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn future_task_get_after_run_and_reset_requires_rerun() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 8); task.run(); task.runAndReset(); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn future_task_executor_shutdown_after_task_completes() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(1); java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 4); pool.submit(task); System.out.println(task.get()); pool.shutdown(); System.out.println(pool.isShutdown());",
    );
    assert_eq!(out, vec!["4", "true"]);
}

#[test]
fn future_task_callable_returns_concatenated_strings() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "vy" + "be"); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn future_task_is_done_false_until_run_called() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "x"); System.out.println(task.isDone()); task.run(); System.out.println(task.isDone());"#,
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn future_task_callable_with_conditional_return() {
    let out = run_main(
        "java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> 5 > 3 ? \"yes\" : \"no\"); task.run(); System.out.println(task.get());",
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn future_task_run_in_thread_pool_via_submit() {
    let out = run_main(
        "java.util.concurrent.ExecutorService pool = java.util.concurrent.Executors.newFixedThreadPool(2); java.util.concurrent.FutureTask<Integer> t1 = new java.util.concurrent.FutureTask<Integer>(() -> 11); java.util.concurrent.FutureTask<Integer> t2 = new java.util.concurrent.FutureTask<Integer>(() -> 22); pool.submit(t1); pool.submit(t2); System.out.println(t1.get() + t2.get()); pool.shutdown();",
    );
    assert_eq!(out, vec!["33"]);
}

#[test]
fn future_task_cancel_after_get_still_returns_result() {
    let out = run_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> 6); task.run(); int v = task.get(); task.cancel(true); System.out.println(v);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn future_task_callable_throws_checked_exception_wrapped() {
    let out = run_in_main(
        "java.util.concurrent.FutureTask<Integer> task = new java.util.concurrent.FutureTask<Integer>(() -> { throw new Exception(\"checked\"); }); task.run(); try { task.get(); System.out.println(\"ok\"); } catch (Exception e) { System.out.println(e.getCause().getMessage()); }",
        "",
    );
    assert_eq!(out, vec!["checked"]);
}

#[test]
fn future_task_runnable_sets_result_after_body() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> System.out.println("body"), "result"); task.run(); System.out.println(task.get());"#,
    );
    assert_eq!(out, vec!["body", "result"]);
}

#[test]
fn future_task_not_cancelled_after_normal_completion() {
    let out = run_main(
        r#"java.util.concurrent.FutureTask<String> task = new java.util.concurrent.FutureTask<String>(() -> "ok"); task.run(); System.out.println(task.isCancelled()); System.out.println(task.isDone());"#,
    );
    assert_eq!(out, vec!["false", "true"]);
}
