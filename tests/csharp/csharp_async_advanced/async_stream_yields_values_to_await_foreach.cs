// vybe-test: csharp/csharp_async_advanced/async_stream_yields_values_to_await_foreach
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

async System.Collections.Generic.IAsyncEnumerable<int> Sequence(){
    for(int i=1;i<=3;i++){
        await System.Threading.Tasks.Task.Yield();
        yield return i;
    }
}
int sum=0;
await foreach(var n in Sequence()) sum+=n;
Console.WriteLine(sum);
