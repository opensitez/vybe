// vybe-test: csharp/csharp_tuples_ranges/conditional_expression
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
__Check((x > 0 ? "positive" : "non-positive").ToString(), "positive");
__Check((x > 10 ? "big" : "small").ToString(), "small");
