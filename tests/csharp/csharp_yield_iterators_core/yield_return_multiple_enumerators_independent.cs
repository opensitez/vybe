// vybe-test: csharp/csharp_yield_iterators_core/yield_return_multiple_enumerators_independent
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Gen(){yield return 1;yield return 2;}
var a=Gen(); var b=Gen(); __Check((a.First()+b.First()).ToString(), "2");
