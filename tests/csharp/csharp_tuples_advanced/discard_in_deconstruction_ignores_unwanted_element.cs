// vybe-test: csharp/csharp_tuples_advanced/discard_in_deconstruction_ignores_unwanted_element
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (first, _, third) = (1, 2, 3);
__Check((first).ToString(), "1"); __Check((third).ToString(), "3");
