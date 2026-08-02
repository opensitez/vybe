// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_write_then_read_returns_same_bytes
// origin: languages/csharp/tests/csharp/test_csharp_memory_stream_roundtrip.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.IO;
var stream = new MemoryStream();
var writer = new StreamWriter(stream);
writer.Write("payload");
writer.Flush();
stream.Position = 0;
var reader = new StreamReader(stream);
__Check((reader.ReadToEnd()).ToString(), "payload");
