// vybe-test: csharp/csharp_yield_iterators_core/yield_return_select_where_pipeline
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> N(){for(int i=0;i<6;i++)yield return i;}
Console.WriteLine(N().Where(x=>x%2==0).Select(x=>x*10).Sum());
