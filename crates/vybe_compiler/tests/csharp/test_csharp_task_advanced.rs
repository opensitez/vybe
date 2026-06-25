//! Advanced `Task` patterns: `Task.Run`, `CancellationToken`, `WhenAny`, `ConfigureAwait`.
use super::helpers::run_csharp;

#[test]
fn task_run_executes_lambda_on_thread_pool() {
    assert_eq!(
        run_csharp(r#"var t=System.Threading.Tasks.Task.Run(()=>42);
Console.WriteLine(t.Result);"#),
        &["42"]
    );
}

#[test]
fn cancellation_token_is_cancelled_after_cancel_called() {
    assert_eq!(
        run_csharp(r#"var cts=new System.Threading.CancellationTokenSource();
cts.Cancel();
Console.WriteLine(cts.Token.IsCancellationRequested);"#),
        &["True"]
    );
}

#[test]
fn task_cancelled_when_token_is_cancelled_before_await() {
    assert_eq!(
        run_csharp(r#"var cts=new System.Threading.CancellationTokenSource();
cts.Cancel();
string result="ok";
try{System.Threading.Tasks.Task.Delay(1000,cts.Token).Wait();}
catch(System.AggregateException){result="cancelled";}
Console.WriteLine(result);"#),
        &["cancelled"]
    );
}

#[test]
fn when_any_returns_first_completed_task() {
    assert_eq!(
        run_csharp(r#"async System.Threading.Tasks.Task<int> Fast()=>await System.Threading.Tasks.Task.FromResult(1);
async System.Threading.Tasks.Task<int> Slow(){await System.Threading.Tasks.Task.Delay(1000);return 2;}
var winner=await System.Threading.Tasks.Task.WhenAny(Fast(),Slow());
Console.WriteLine(winner.Result);"#),
        &["1"]
    );
}

#[test]
fn task_continuation_runs_after_completion() {
    assert_eq!(
        run_csharp(r#"int result=0;
System.Threading.Tasks.Task.Run(()=>7)
    .ContinueWith(t=>result=t.Result*2)
    .Wait();
Console.WriteLine(result);"#),
        &["14"]
    );
}

#[test]
fn task_from_exception_rethrows_on_await() {
    assert_eq!(
        run_csharp(r#"string msg="";
var t=System.Threading.Tasks.Task.FromException(new System.Exception("boom"));
try{t.Wait();}catch(System.AggregateException ae){msg=ae.InnerException.Message;}
Console.WriteLine(msg);"#),
        &["boom"]
    );
}

#[test]
fn task_completed_task_already_in_ran_to_completion_state() {
    assert_eq!(
        run_csharp(r#"var t=System.Threading.Tasks.Task.CompletedTask;
Console.WriteLine(t.IsCompleted);"#),
        &["True"]
    );
}
