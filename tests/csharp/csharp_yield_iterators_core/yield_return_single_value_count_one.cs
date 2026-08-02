// vybe-test: csharp/csharp_yield_iterators_core/yield_return_single_value_count_one
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> One(){yield return 42;}
int c=0; foreach(var _ in One()) c++; Console.WriteLine(c);
