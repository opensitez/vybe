// vybe-test: csharp/csharp_yield_iterators_core/yield_break_after_zero_yields_empty
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Gen(){yield break;yield return 1;}
__Check((Gen().Count()).ToString(), "0");
