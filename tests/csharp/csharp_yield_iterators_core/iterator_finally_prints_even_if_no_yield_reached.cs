// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_prints_even_if_no_yield_reached
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(bool ok){try{if(!ok)yield break;yield return 1;}finally{Console.WriteLine("end");}}
foreach(var _ in Gen(false)){}
