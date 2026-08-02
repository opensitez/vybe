// vybe-test: csharp/csharp_iasync_enumerable/cancellation_after_three_yields_partial_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 10; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    int count = 0;
    try {
        await foreach (var x in Stream(cts.Token)) {
            count++;
            if (count == 3) cts.Cancel();
        }
    } catch (System.OperationCanceledException) { }
    Console.WriteLine(count);
}
Run().Wait();
