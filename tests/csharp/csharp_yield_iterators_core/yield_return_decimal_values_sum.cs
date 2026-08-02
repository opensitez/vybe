// vybe-test: csharp/csharp_yield_iterators_core/yield_return_decimal_values_sum
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<decimal> D(){yield return 1.5m;yield return 2.5m;}
__Check((D().Sum()).ToString(), "4.0");
