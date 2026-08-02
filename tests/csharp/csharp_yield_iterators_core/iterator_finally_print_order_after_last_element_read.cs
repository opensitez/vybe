// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_print_order_after_last_element_read
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){try{yield return 10;yield return 20;}finally{Console.WriteLine("after");}}
foreach(var n in Gen()) Console.WriteLine(n);
