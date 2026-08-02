// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_static_method
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Seq{public static System.Collections.Generic.IEnumerable<int> Twice(int n){yield return n;yield return n*2;}}
__Check((string.Join(",",Seq.Twice(5))).ToString(), "5,10");
