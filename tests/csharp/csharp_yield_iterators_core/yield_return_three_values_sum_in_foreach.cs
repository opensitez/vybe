// vybe-test: csharp/csharp_yield_iterators_core/yield_return_three_values_sum_in_foreach
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Gen(){yield return 1;yield return 2;yield return 3;}
int s=0; foreach(var n in Gen()) s+=n; Console.WriteLine(s);
