// vybe-test: csharp/csharp_linq_numeric/long_count_works_on_large_range
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long c=Enumerable.Range(0,1000).LongCount();
__Check((c).ToString(), "1000");
