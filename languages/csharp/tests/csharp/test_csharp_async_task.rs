//! `Task`, `Task<T>`, async/await scheduling and result propagation.
use super::helpers::run_csharp;

#[test]
fn task_from_result_completes_synchronously_with_value() {
    assert_eq!(
        run_csharp(
            r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.FromResult(42);
    Console.WriteLine(await t);
}
Run().Wait();
"#
        ),
        &["42"]
    );
}

#[test]
fn awaiting_task_delay_resumes_execution_after_pause() {
    assert_eq!(
        run_csharp(
            r#"
async System.Threading.Tasks.Task Run() {
    Console.WriteLine("before");
    await System.Threading.Tasks.Task.Delay(1);
    Console.WriteLine("after");
}
Run().Wait();
"#
        ),
        &["before", "after"]
    );
}

#[test]
fn async_method_returns_task_t_with_computed_value() {
    assert_eq!(
        run_csharp(
            r#"
async System.Threading.Tasks.Task<int> Compute() {
    await System.Threading.Tasks.Task.Yield();
    return 7;
}
Console.WriteLine(Compute().Result);
"#
        ),
        &["7"]
    );
}

#[test]
fn exception_in_async_method_propagates_through_await() {
    assert_eq!(
        run_csharp(
            r#"
async System.Threading.Tasks.Task Fail() {
    await System.Threading.Tasks.Task.Yield();
    throw new System.Exception("async fail");
}
string msg = "";
try { Fail().Wait(); }
catch (System.AggregateException ae) { msg = ae.InnerException.Message; }
Console.WriteLine(msg);
"#
        ),
        &["async fail"]
    );
}

#[test]
fn task_when_all_waits_for_multiple_tasks_to_complete() {
    assert_eq!(
        run_csharp(
            r#"
async System.Threading.Tasks.Task<int> Val(int n) {
    await System.Threading.Tasks.Task.Yield();
    return n;
}
var results = System.Threading.Tasks.Task.WhenAll(Val(1), Val(2), Val(3)).Result;
int sum = 0;
foreach (var r in results) sum += r;
Console.WriteLine(sum);
"#
        ),
        &["6"]
    );
}
