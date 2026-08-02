// vybe-test: csharp/csharp_yield_iterators_core/yield_return_in_while_produces_infinite_prefix
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Up(){int n=0;while(n<4){yield return n;n++;}}
Console.WriteLine(string.Join(",",Up()));
