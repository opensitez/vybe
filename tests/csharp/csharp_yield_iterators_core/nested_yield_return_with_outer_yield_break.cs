// vybe-test: csharp/csharp_yield_iterators_core/nested_yield_return_with_outer_yield_break
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in new[]{1,2,3}){if(x==2)yield break;yield return x;}}
Console.WriteLine(string.Join(",",Outer()));
