// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Outer(){System.Collections.Generic.IEnumerable<int> Inner(){yield return 3;} foreach(var x in Inner())yield return x;}
Console.WriteLine(Outer().First());
