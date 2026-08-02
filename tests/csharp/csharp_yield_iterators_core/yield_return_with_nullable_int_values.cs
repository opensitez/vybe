// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_nullable_int_values
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int?> Maybe(){yield return null;yield return 4;}
__Check((string.Join(",",Maybe().Select(x=>x??0))).ToString(), "0,4");
