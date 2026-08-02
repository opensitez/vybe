// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_string_length_total_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<string> Stream() {
    yield return "ab";
    yield return "cde";
    yield return "f";
}
async System.Threading.Tasks.Task Run() {
    int totalLen = 0;
    await foreach (var s in Stream()) totalLen += s.Length;
    Console.WriteLine(totalLen);
}
Run().Wait();
