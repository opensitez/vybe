// vybe-test: csharp/csharp_iasync_enumerable/params_style_loop_yield_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<int> FromValues(int a, int b, int c, int d) {
    yield return a;
    yield return b;
    yield return c;
    yield return d;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in FromValues(1, 2, 3, 4)) count++;
    Console.WriteLine(count);
}
Run().Wait();
