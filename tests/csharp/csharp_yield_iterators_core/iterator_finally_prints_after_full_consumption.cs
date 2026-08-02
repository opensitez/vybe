// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_prints_after_full_consumption
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;yield return 2;}finally{Console.WriteLine("fin");}}
foreach(var _ in Gen()){}
