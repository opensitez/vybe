// vybe-test: csharp/csharp_yield_iterators_core/yield_break_at_start_before_any_yield
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Gen(){yield break;yield return 1;}
__Check((Gen().Any()).ToString(), "False");
