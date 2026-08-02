// vybe-test: csharp/csharp_tuples_advanced/tuple_deconstruction_assigns_to_separate_locals
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (a, b, c) = (10, 20, 30);
__Check((a+b+c).ToString(), "60");
