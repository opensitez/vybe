// vybe-test: csharp/csharp_yield_iterators_core/nested_yield_with_conditional_inner_skip
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Inner(bool ok){if(ok)yield return 9;}
System.Collections.Generic.IEnumerable<int> Outer(bool ok){foreach(var x in Inner(ok))yield return x;yield return 1;}
Console.WriteLine(string.Join(",",Outer(false)));
