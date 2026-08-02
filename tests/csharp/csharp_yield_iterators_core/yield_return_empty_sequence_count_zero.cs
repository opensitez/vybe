// vybe-test: csharp/csharp_yield_iterators_core/yield_return_empty_sequence_count_zero
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Empty(){yield break;}
int c=0; foreach(var _ in Empty()) c++; Console.WriteLine(c);
