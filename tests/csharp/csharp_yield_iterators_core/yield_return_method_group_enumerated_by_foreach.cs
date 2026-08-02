// vybe-test: csharp/csharp_yield_iterators_core/yield_return_method_group_enumerated_by_foreach
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Range(int n){for(int i=0;i<n;i++)yield return i;}
void Run(System.Collections.Generic.IEnumerable<int> src){Console.WriteLine(src.Sum());}
Run(Range(5));
