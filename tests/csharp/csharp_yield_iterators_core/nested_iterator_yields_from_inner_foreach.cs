// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_yields_from_inner_foreach
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Inner(){yield return 1;yield return 2;}
System.Collections.Generic.IEnumerable<int> Outer(){foreach(var x in Inner())yield return x;foreach(var x in Inner())yield return x;}
Console.WriteLine(string.Join(",",Outer()));
