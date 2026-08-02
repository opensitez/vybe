// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_count_matches_flat_length
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> A(){yield return 1;yield return 2;}
System.Collections.Generic.IEnumerable<int> B(){foreach(var x in A())yield return x;}
Console.WriteLine(B().Count());
