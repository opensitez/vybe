// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_parameterized_start_index
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> From(int start,int count){for(int i=0;i<count;i++)yield return start+i;}
Console.WriteLine(string.Join(",",From(5,3)));
