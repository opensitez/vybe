// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_not_run_until_started_iteration
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

int fin=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{fin=1;Console.WriteLine(fin);}}
var seq=Gen(); Console.WriteLine(fin); foreach(var _ in seq){}
