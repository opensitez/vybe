// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_seek_begin_repositions_read_cursor_for_second_pass
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
writer.Write("ab");
writer.Flush();
stream.Seek(0, SeekOrigin.Begin);
__Check((stream.ReadByte()).ToString(), "97");
stream.Seek(0, SeekOrigin.Begin);
__Check((stream.ReadByte()).ToString(), "97");
