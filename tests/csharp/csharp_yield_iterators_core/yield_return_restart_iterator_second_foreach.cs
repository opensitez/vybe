// vybe-test: csharp/csharp_yield_iterators_core/yield_return_restart_iterator_second_foreach
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Two(){yield return 10;yield return 20;}
int sum=0; foreach(var n in Two()) sum+=n; foreach(var n in Two()) sum+=n; Console.WriteLine(sum);
