// vybe-test: csharp/csharp_iasync_enumerable/generic_async_enumerable_method_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<T> Repeat<T>(T value, int times) {
    for (int i = 0; i < times; i++) yield return value;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Repeat(7, 4)) count++;
    Console.WriteLine(count);
}
Run().Wait();
