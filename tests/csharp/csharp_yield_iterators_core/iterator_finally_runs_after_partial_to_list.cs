// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_runs_after_partial_to_list
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;yield return 2;yield return 3;}finally{__Check(("close").ToString(), "close");}}
Gen().Take(2).ToList();
