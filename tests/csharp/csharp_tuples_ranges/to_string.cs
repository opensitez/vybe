// vybe-test: csharp/csharp_tuples_ranges/to_string
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 42;
__Check((x.ToString()).ToString(), "42");
__Check((42.ToString()).ToString(), "42");
