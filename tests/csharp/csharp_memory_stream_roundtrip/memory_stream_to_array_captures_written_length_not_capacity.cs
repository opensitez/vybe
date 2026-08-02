// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_to_array_captures_written_length_not_capacity
// origin: languages/csharp/tests/csharp/test_csharp_memory_stream_roundtrip.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.IO;
var stream = new MemoryStream();
stream.WriteByte(1);
stream.WriteByte(2);
var bytes = stream.ToArray();
__Check((bytes.Length).ToString(), "2");
__Check((bytes[1]).ToString(), "2");
