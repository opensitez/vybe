// vybe-test: csharp/csharp_yield_iterators_core/yield_return_nested_try_finally_inner_finally_print
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){try{try{yield return 1;}finally{Console.WriteLine("inner");}}finally{Console.WriteLine("outer");}}
foreach(var _ in Gen()){}
