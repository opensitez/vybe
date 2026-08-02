// vybe-test: csharp/csharp_yield_iterators_core/iterator_finally_prints_when_consumer_breaks_early
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){try{for(int i=0;i<5;i++)yield return i;}finally{Console.WriteLine("done");}}
foreach(var n in Gen()){if(n==2)break;}
