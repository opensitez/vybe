// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_explicit_ienumerable_interface
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Nums:System.Collections.Generic.IEnumerable<int>{public System.Collections.Generic.IEnumerator<int> GetEnumerator(){yield return 2;yield return 4;}System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();}
__Check((new Nums().Sum()).ToString(), "6");
