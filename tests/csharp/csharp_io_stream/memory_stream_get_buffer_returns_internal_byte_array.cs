// vybe-test: csharp/csharp_io_stream/memory_stream_get_buffer_returns_internal_byte_array
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var ms = new System.IO.MemoryStream(new byte[]{1,2,3});
__Check((ms.Length).ToString(), "3");
