// vybe-test: csharp/csharp_io_stream/string_reader_reads_line_from_in_memory_string
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var sr = new System.IO.StringReader("line one\nline two");
__Check((sr.ReadLine()).ToString(), "line one");
