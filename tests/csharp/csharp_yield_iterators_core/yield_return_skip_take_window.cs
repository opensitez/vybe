// vybe-test: csharp/csharp_yield_iterators_core/yield_return_skip_take_window
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> N(){for(int i=0;i<10;i++)yield return i;}
Console.WriteLine(string.Join(",",N().Skip(3).Take(2)));
