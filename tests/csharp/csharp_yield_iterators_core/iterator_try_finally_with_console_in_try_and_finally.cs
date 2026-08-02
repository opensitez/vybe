// vybe-test: csharp/csharp_yield_iterators_core/iterator_try_finally_with_console_in_try_and_finally
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){try{Console.WriteLine("try");yield return 5;}finally{Console.WriteLine("finally");}}
foreach(var n in Gen()) Console.WriteLine(n);
