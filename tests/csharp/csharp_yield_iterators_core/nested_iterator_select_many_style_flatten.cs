// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_select_many_style_flatten
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Pair(int n){yield return n;yield return n+1;}
__Check((string.Join(",",new[]{1,2}.SelectMany(Pair))).ToString(), "1,2,2,3");
