// vybe-test: csharp/csharp_iasync_enumerable/async_enumerable_char_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<char> Chars() {
    yield return 'a';
    yield return 'b';
    yield return 'c';
    yield return 'd';
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var ch in Chars()) count++;
    Console.WriteLine(count);
}
Run().Wait();
