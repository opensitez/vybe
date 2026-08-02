// vybe-test: csharp/csharp_iasync_enumerable/skip_first_two_via_flag_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 0; i < 6; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int seen = 0;
    int count = 0;
    await foreach (var x in Stream()) {
        seen++;
        if (seen <= 2) continue;
        count++;
    }
    Console.WriteLine(count);
}
Run().Wait();
