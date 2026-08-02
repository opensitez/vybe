// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_prints_after_yield_break
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;yield break;yield return 9;}finally{Console.WriteLine("cleanup");}}
foreach(var _ in Gen()){}
