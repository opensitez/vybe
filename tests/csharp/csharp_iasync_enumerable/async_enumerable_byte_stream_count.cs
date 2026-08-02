// vybe-test: csharp/csharp_iasync_enumerable/async_enumerable_byte_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

async System.Collections.Generic.IAsyncEnumerable<byte> Bytes() {
    yield return (byte)1;
    yield return (byte)2;
    yield return (byte)3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var b in Bytes()) count++;
    Console.WriteLine(count);
}
Run().Wait();
