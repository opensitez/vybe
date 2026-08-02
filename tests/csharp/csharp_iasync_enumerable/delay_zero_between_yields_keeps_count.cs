// vybe-test: csharp/csharp_iasync_enumerable/delay_zero_between_yields_keeps_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    await System.Threading.Tasks.Task.Delay(0);
    yield return 2;
    await System.Threading.Tasks.Task.Delay(0);
    yield return 3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
