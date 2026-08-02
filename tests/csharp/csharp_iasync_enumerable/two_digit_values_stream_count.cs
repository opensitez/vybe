// vybe-test: csharp/csharp_iasync_enumerable/two_digit_values_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Tens() {
    yield return 10;
    yield return 20;
    yield return 30;
    yield return 40;
    yield return 50;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Tens()) count++;
    Console.WriteLine(count);
}
Run().Wait();
