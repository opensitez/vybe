// vybe-test: csharp/csharp_checked_unchecked/default_arithmetic_is_unchecked_for_performance
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=int.MaxValue; x++;
__Check((x==int.MinValue).ToString(), "True");
