// vybe-test: csharp/csharp_iasync_enumerable/static_async_enumerable_method_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

class Factory {
    public static async System.Collections.Generic.IAsyncEnumerable<int> Three() {
        yield return 1;
        yield return 2;
        yield return 3;
    }
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Factory.Three()) count++;
    Console.WriteLine(count);
}
Run().Wait();
