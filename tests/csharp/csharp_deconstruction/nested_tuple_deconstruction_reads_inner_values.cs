// vybe-test: csharp/csharp_deconstruction/nested_tuple_deconstruction_reads_inner_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ((x, y), label) = ((5, 6), "pt");
__Check((label).ToString(), "pt");
__Check((x + y).ToString(), "11");
