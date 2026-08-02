// vybe-test: csharp/csharp_yield_iterators_core/yield_return_boxed_objects_via_enumerable
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<object> Box(){yield return 1;yield return "two";}
int c=0; foreach(var _ in Box()) c++; Console.WriteLine(c);
