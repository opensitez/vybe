// vybe-test: csharp/csharp_iasync_enumerable/async_local_function_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Threading.Tasks.Task Run() {
    async System.Collections.Generic.IAsyncEnumerable<int> Local() {
        for (int i = 0; i < 5; i++) yield return i;
    }
    int count = 0;
    await foreach (var x in Local()) count++;
    Console.WriteLine(count);
}
Run().Wait();
