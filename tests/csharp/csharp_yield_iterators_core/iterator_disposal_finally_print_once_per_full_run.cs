// vybe-test: csharp/csharp_yield_iterators_core/iterator_disposal_finally_print_once_per_full_run
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

int hits=0; System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 1;}finally{hits++;Console.WriteLine(hits);}}
foreach(var _ in Gen()){} Console.WriteLine(hits);
