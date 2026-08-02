// vybe-test: csharp/csharp_yield_iterators_core/nested_three_level_iterator_chain
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> A(){yield return 1;}
System.Collections.Generic.IEnumerable<int> B(){foreach(var x in A())yield return x+10;}
System.Collections.Generic.IEnumerable<int> C(){foreach(var x in B())yield return x+100;}
Console.WriteLine(string.Join(",",C()));
