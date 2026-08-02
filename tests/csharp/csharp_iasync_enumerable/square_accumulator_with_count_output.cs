// vybe-test: csharp/csharp_iasync_enumerable/square_accumulator_with_count_output
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 2;
    yield return 3;
    yield return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    int sumSq = 0;
    await foreach (var x in Stream()) { sumSq += x * x; count++; }
    Console.WriteLine(count);
}
Run().Wait();
