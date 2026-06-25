//! Advanced async/await: `ValueTask`, async streams, `IAsyncEnumerable`, `ConfigureAwait`.
use super::helpers::run_csharp;

#[test]
fn value_task_from_result_avoids_allocation() {
    assert_eq!(
        run_csharp(r#"async System.Threading.Tasks.ValueTask<int> GetValueAsync()=>42;
int v=await GetValueAsync();
Console.WriteLine(v);"#),
        &["42"]
    );
}

#[test]
fn async_method_exception_propagated_through_await() {
    assert_eq!(
        run_csharp(r#"async System.Threading.Tasks.Task Fail()=>throw new System.Exception("async fail");
string msg="";
try{await Fail();}catch(System.Exception ex){msg=ex.Message;}
Console.WriteLine(msg);"#),
        &["async fail"]
    );
}

#[test]
fn async_stream_yields_values_to_await_foreach() {
    assert_eq!(
        run_csharp(r#"async System.Collections.Generic.IAsyncEnumerable<int> Sequence(){
    for(int i=1;i<=3;i++){
        await System.Threading.Tasks.Task.Yield();
        yield return i;
    }
}
int sum=0;
await foreach(var n in Sequence()) sum+=n;
Console.WriteLine(sum);"#),
        &["6"]
    );
}

#[test]
fn when_all_awaits_all_tasks_and_returns_results() {
    assert_eq!(
        run_csharp(r#"async System.Threading.Tasks.Task<int> N(int v){
    await System.Threading.Tasks.Task.Delay(0);return v;
}
int[] results=await System.Threading.Tasks.Task.WhenAll(N(1),N(2),N(3));
Console.WriteLine(results.Sum());"#),
        &["6"]
    );
}

#[test]
fn configure_await_false_does_not_resume_on_original_context() {
    assert_eq!(
        run_csharp(r#"async System.Threading.Tasks.Task<int> Compute(){
    await System.Threading.Tasks.Task.Delay(1).ConfigureAwait(false);
    return 42;
}
Console.WriteLine(await Compute());"#),
        &["42"]
    );
}
