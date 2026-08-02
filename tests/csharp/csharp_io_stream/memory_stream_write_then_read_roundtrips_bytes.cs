// vybe-test: csharp/csharp_io_stream/memory_stream_write_then_read_roundtrips_bytes
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var ms = new System.IO.MemoryStream();
ms.WriteByte(42);
ms.Position = 0;
__Check((ms.ReadByte()).ToString(), "42");
