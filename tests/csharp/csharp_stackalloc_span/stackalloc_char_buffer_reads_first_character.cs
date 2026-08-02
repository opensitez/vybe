// vybe-test: csharp/csharp_stackalloc_span/stackalloc_char_buffer_reads_first_character
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<char> buf=stackalloc char[3]{'a','b','c'}; __Check((buf[0]).ToString(), "a");
