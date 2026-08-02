// vybe-test: csharp/csharp_yield_iterators_core/yield_return_from_instance_method_on_class
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

class Counter{public System.Collections.Generic.IEnumerable<int> Range(int n){for(int i=0;i<n;i++)yield return i;}}
Console.WriteLine(new Counter().Range(4).Sum());
