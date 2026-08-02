// vybe-test: csharp/csharp_io_stream/stream_seek_repositions_for_re_read
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var ms = new System.IO.MemoryStream();
ms.WriteByte(7);
ms.Seek(0, System.IO.SeekOrigin.Begin);
__Check((ms.ReadByte()).ToString(), "7");
