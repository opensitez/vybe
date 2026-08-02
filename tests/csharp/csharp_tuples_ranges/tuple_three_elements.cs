// vybe-test: csharp/csharp_tuples_ranges/tuple_three_elements
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t = (1, 2, 3);
__Check((t.Item1 + t.Item2 + t.Item3).ToString(), "6");
