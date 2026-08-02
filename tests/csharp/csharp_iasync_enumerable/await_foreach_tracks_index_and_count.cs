// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_tracks_index_and_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 5;
    yield return 10;
    yield return 15;
}
async System.Threading.Tasks.Task Run() {
    int index = 0;
    await foreach (var x in Stream()) index++;
    Console.WriteLine(index);
}
Run().Wait();
