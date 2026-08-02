// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_runs_once_per_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

int fin=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{fin++;Console.WriteLine(fin);}}
foreach(var _ in Gen()){} foreach(var _ in Gen()){}
