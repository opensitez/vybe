// vybe-test: csharp/csharp_linq_aggregate_element/min_by_first_of_tie
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{"x","y","z"}.MinBy(s=>s.Length)).ToString(), "x");
