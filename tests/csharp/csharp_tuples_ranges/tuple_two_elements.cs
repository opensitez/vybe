// vybe-test: csharp/csharp_tuples_ranges/tuple_two_elements
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t = (10, "hello");
__Check((t.Item1).ToString(), "10");
__Check((t.Item2).ToString(), "hello");
