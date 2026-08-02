// vybe-test: csharp/csharp_yield_iterators_core/yield_return_repeated_value_pattern
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Repeat(int v,int n){for(int i=0;i<n;i++)yield return v;}
Console.WriteLine(string.Join(",",Repeat(7,3)));
