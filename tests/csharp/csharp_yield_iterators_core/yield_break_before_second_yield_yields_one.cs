// vybe-test: csharp/csharp_yield_iterators_core/yield_break_before_second_yield_yields_one
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){yield return 1;yield break;yield return 2;}
int c=0; foreach(var _ in Gen()) c++; Console.WriteLine(c);
