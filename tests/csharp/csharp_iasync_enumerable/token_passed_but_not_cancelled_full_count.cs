// vybe-test: csharp/csharp_iasync_enumerable/token_passed_but_not_cancelled_full_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 9; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    int count = 0;
    await foreach (var x in Stream(cts.Token)) count++;
    Console.WriteLine(count);
}
Run().Wait();
