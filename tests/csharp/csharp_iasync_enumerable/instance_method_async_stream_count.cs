// vybe-test: csharp/csharp_iasync_enumerable/instance_method_async_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

class Counter {
    public async System.Collections.Generic.IAsyncEnumerable<int> Stream(int n) {
        for (int i = 0; i < n; i++) yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var c = new Counter();
    int count = 0;
    await foreach (var x in c.Stream(7)) count++;
    Console.WriteLine(count);
}
Run().Wait();
