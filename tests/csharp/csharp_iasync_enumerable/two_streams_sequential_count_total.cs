// vybe-test: csharp/csharp_iasync_enumerable/two_streams_sequential_count_total
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> A() {
    yield return 1;
    yield return 2;
}
async System.Collections.Generic.IAsyncEnumerable<int> B() {
    yield return 10;
    yield return 20;
    yield return 30;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in A()) count++;
    await foreach (var x in B()) count++;
    Console.WriteLine(count);
}
Run().Wait();
