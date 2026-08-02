// vybe-test: csharp/csharp_iasync_enumerable/cancellation_token_none_runs_entire_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken cancellationToken) {
    for (int i = 0; i < 4; i++) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream(System.Threading.CancellationToken.None)) count++;
    Console.WriteLine(count);
}
Run().Wait();
