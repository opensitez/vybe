// vybe-test: csharp/csharp_yield_iterators_core/yield_return_dispose_pattern_prints_on_iterator_close
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Track(){try{yield return 1;}finally{__Check(("dispose").ToString(), "1");}}
using var e=Track().GetEnumerator(); e.MoveNext(); __Check((e.Current).ToString(), "dispose");
