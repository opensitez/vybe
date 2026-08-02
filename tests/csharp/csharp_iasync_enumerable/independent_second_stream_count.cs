// vybe-test: csharp/csharp_iasync_enumerable/independent_second_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Pair() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Pair()) count++;
    Console.WriteLine(count);
}
Run().Wait();
