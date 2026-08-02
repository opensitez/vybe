// vybe-test: csharp/csharp_io_stream/binary_writer_reader_roundtrips_int32
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var ms = new System.IO.MemoryStream();
using(var bw = new System.IO.BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen:true))
    bw.Write(12345);
ms.Position = 0;
using var br = new System.IO.BinaryReader(ms);
__Check((br.ReadInt32()).ToString(), "12345");
