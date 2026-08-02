// vybe-test: csharp/csharp_yield_iterators_core/nested_iterator_yields_concatenated_strings
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<string> Words(){yield return "a";yield return "b";}
System.Collections.Generic.IEnumerable<string> Twice(){foreach(var w in Words())yield return w;foreach(var w in Words())yield return w;}
Console.WriteLine(string.Join("",Twice()));
