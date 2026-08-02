// vybe-test: csharp/csharp_io_stream/stream_writer_reader_roundtrip_text_line
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var ms = new System.IO.MemoryStream();
using(var sw = new System.IO.StreamWriter(ms, leaveOpen:true)) sw.WriteLine("hello");
ms.Position = 0;
using var sr = new System.IO.StreamReader(ms);
__Check((sr.ReadLine()).ToString(), "hello");
