// vybe-test: csharp/csharp_yield_iterators_core/nested_yield_return_with_inner_yield_break
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Inner(){yield return 1;yield break;yield return 9;}
System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in Inner())yield return x;yield return 2;}
Console.WriteLine(string.Join(",",Outer()));
