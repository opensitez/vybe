// vybe-test: csharp/csharp_yield_iterators_core/yield_return_infinite_prefix_take_three
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Naturals(){int n=0;while(true)yield return n++;}
Console.WriteLine(string.Join(",",Naturals().Take(3)));
